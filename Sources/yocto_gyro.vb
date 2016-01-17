'*********************************************************************
'*
'* $Id: yocto_gyro.vb 22698 2016-01-12 23:15:02Z seb $
'*
'* Implements yFindGyro(), the high-level API for Gyro functions
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


Module yocto_gyro

  REM --- (generated code: YQt return codes)
    REM --- (end of generated code: YQt return codes)
  REM --- (generated code: YQt globals)

  Public Delegate Sub YQtValueCallback(ByVal func As YQt, ByVal value As String)
  Public Delegate Sub YQtTimedReportCallback(ByVal func As YQt, ByVal measure As YMeasure)
  REM --- (end of generated code: YQt globals)

  REM --- (generated code: YQt class start)

  '''*
  ''' <summary>
  '''   The Yoctopuce API YQt class provides direct access to the Yocto3D attitude estimation
  '''   using a quaternion.
  ''' <para>
  '''   It is usually not needed to use the YQt class directly, as the
  '''   YGyro class provides a more convenient higher-level interface.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YQt
    Inherits YSensor
    REM --- (end of generated code: YQt class start)

    REM --- (generated code: YQt definitions)
    REM --- (end of generated code: YQt definitions)

    REM --- (generated code: YQt attributes declaration)
    Protected _valueCallbackQt As YQtValueCallback
    Protected _timedReportCallbackQt As YQtTimedReportCallback
    REM --- (end of generated code: YQt attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Qt"
      REM --- (generated code: YQt attributes initialization)
      _valueCallbackQt = Nothing
      _timedReportCallbackQt = Nothing
      REM --- (end of generated code: YQt attributes initialization)
    End Sub

  REM --- (generated code: YQt private methods declaration)

    Protected Overrides Function _parseAttr(ByRef member As TJSONRECORD) As Integer
      Return MyBase._parseAttr(member)
    End Function

    REM --- (end of generated code: YQt private methods declaration)

    REM --- (generated code: YQt public methods declaration)
    '''*
    ''' <summary>
    '''   Retrieves a quaternion component for a given identifier.
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
    '''   This function does not require that the quaternion component is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YQt.isOnline()</c> to test if the quaternion component is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a quaternion component by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the quaternion component
    ''' </param>
    ''' <returns>
    '''   a <c>YQt</c> object allowing you to drive the quaternion component.
    ''' </returns>
    '''/
    Public Shared Function FindQt(func As String) As YQt
      Dim obj As YQt
      obj = CType(YFunction._FindFromCache("Qt", func), YQt)
      If ((obj Is Nothing)) Then
        obj = New YQt(func)
        YFunction._AddToCache("Qt", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YQtValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackQt = callback
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
      If (Not (Me._valueCallbackQt Is Nothing)) Then
        Me._valueCallbackQt(Me, value)
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
    Public Overloads Function registerTimedReportCallback(callback As YQtTimedReportCallback) As Integer
      Dim sensor As YSensor
      sensor = Me
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(sensor, True)
      Else
        YFunction._UpdateTimedReportCallbackList(sensor, False)
      End If
      Me._timedReportCallbackQt = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackQt Is Nothing)) Then
        Me._timedReportCallbackQt(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of quaternion components started using <c>yFirstQt()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YQt</c> object, corresponding to
    '''   a quaternion component currently online, or a <c>null</c> pointer
    '''   if there are no more quaternion components to enumerate.
    ''' </returns>
    '''/
    Public Function nextQt() As YQt
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YQt.FindQt(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of quaternion components currently accessible.
    ''' <para>
    '''   Use the method <c>YQt.nextQt()</c> to iterate on
    '''   next quaternion components.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YQt</c> object, corresponding to
    '''   the first quaternion component currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstQt() As YQt
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Qt", 0, p, size, neededsize, errmsg)
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
      Return YQt.FindQt(serial + "." + funcId)
    End Function

    REM --- (end of generated code: YQt public methods declaration)

  End Class

  REM --- (generated code: Qt functions)

  '''*
  ''' <summary>
  '''   Retrieves a quaternion component for a given identifier.
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
  '''   This function does not require that the quaternion component is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YQt.isOnline()</c> to test if the quaternion component is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a quaternion component by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the quaternion component
  ''' </param>
  ''' <returns>
  '''   a <c>YQt</c> object allowing you to drive the quaternion component.
  ''' </returns>
  '''/
  Public Function yFindQt(ByVal func As String) As YQt
    Return YQt.FindQt(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of quaternion components currently accessible.
  ''' <para>
  '''   Use the method <c>YQt.nextQt()</c> to iterate on
  '''   next quaternion components.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YQt</c> object, corresponding to
  '''   the first quaternion component currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstQt() As YQt
    Return YQt.FirstQt()
  End Function


  REM --- (end of generated code: Qt functions)

  REM --- (generated code: YGyro return codes)
    REM --- (end of generated code: YGyro return codes)
  REM --- (generated code: YGyro globals)

  Public Const Y_XVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_YVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_ZVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Delegate Sub YGyroValueCallback(ByVal func As YGyro, ByVal value As String)
  Public Delegate Sub YGyroTimedReportCallback(ByVal func As YGyro, ByVal measure As YMeasure)
  REM --- (end of generated code: YGyro globals)
  Public Delegate Sub YQuatCallback(ByVal func As YGyro, ByVal w As Double, ByVal x As Double, ByVal y As Double, ByVal z As Double)
  Public Delegate Sub YAnglesCallback(ByVal func As YGyro, ByVal roll As Double, ByVal pitch As Double, ByVal head As Double)

  Sub yInternalGyroCallback(ByVal obj As YQt, ByVal value As String)
    Dim gyro As YGyro
    Dim tmp As String
    Dim idx As Integer
    Dim dbl_value As Double
    gyro = CType(obj.get_userData(), YGyro)
    If gyro Is Nothing Then
      Return
    End If
    tmp = obj.get_functionId().Substring(2)
    idx = Convert.ToInt32(tmp)
    dbl_value = Convert.ToDouble(value)
    gyro._invokeGyroCallbacks(idx, dbl_value)
  End Sub

  REM --- (generated code: YGyro class start)

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
  Public Class YGyro
    Inherits YSensor
    REM --- (end of generated code: YGyro class start)

    REM --- (generated code: YGyro definitions)
    Public Const XVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const YVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const ZVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
    REM --- (end of generated code: YGyro definitions)

    REM --- (generated code: YGyro attributes declaration)
    Protected _xValue As Double
    Protected _yValue As Double
    Protected _zValue As Double
    Protected _valueCallbackGyro As YGyroValueCallback
    Protected _timedReportCallbackGyro As YGyroTimedReportCallback
    Protected _qt_stamp As Integer
    Protected _qt_w As YQt
    Protected _qt_x As YQt
    Protected _qt_y As YQt
    Protected _qt_z As YQt
    Protected _w As Double
    Protected _x As Double
    Protected _y As Double
    Protected _z As Double
    Protected _angles_stamp As Integer
    Protected _head As Double
    Protected _pitch As Double
    Protected _roll As Double
    Protected _quatCallback As YQuatCallback
    Protected _anglesCallback As YAnglesCallback
    REM --- (end of generated code: YGyro attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Gyro"
      REM --- (generated code: YGyro attributes initialization)
      _xValue = XVALUE_INVALID
      _yValue = YVALUE_INVALID
      _zValue = ZVALUE_INVALID
      _valueCallbackGyro = Nothing
      _timedReportCallbackGyro = Nothing
      _qt_stamp = 0
      _w = 0
      _x = 0
      _y = 0
      _z = 0
      _angles_stamp = 0
      _head = 0
      _pitch = 0
      _roll = 0
      REM --- (end of generated code: YGyro attributes initialization)
    End Sub

    REM --- (generated code: YGyro private methods declaration)

    Protected Overrides Function _parseAttr(ByRef member As TJSONRECORD) As Integer
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
      Return MyBase._parseAttr(member)
    End Function

    REM --- (end of generated code: YGyro private methods declaration)

    REM --- (generated code: YGyro public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the angular velocity around the X axis of the device, as a floating point number.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the angular velocity around the X axis of the device, as a
    '''   floating point number
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
    '''   Returns the angular velocity around the Y axis of the device, as a floating point number.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the angular velocity around the Y axis of the device, as a
    '''   floating point number
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
    '''   Returns the angular velocity around the Z axis of the device, as a floating point number.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the angular velocity around the Z axis of the device, as a
    '''   floating point number
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

    '''*
    ''' <summary>
    '''   Retrieves a gyroscope for a given identifier.
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
    '''   This function does not require that the gyroscope is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YGyro.isOnline()</c> to test if the gyroscope is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a gyroscope by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the gyroscope
    ''' </param>
    ''' <returns>
    '''   a <c>YGyro</c> object allowing you to drive the gyroscope.
    ''' </returns>
    '''/
    Public Shared Function FindGyro(func As String) As YGyro
      Dim obj As YGyro
      obj = CType(YFunction._FindFromCache("Gyro", func), YGyro)
      If ((obj Is Nothing)) Then
        obj = New YGyro(func)
        YFunction._AddToCache("Gyro", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YGyroValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackGyro = callback
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
      If (Not (Me._valueCallbackGyro Is Nothing)) Then
        Me._valueCallbackGyro(Me, value)
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
    Public Overloads Function registerTimedReportCallback(callback As YGyroTimedReportCallback) As Integer
      Dim sensor As YSensor
      sensor = Me
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(sensor, True)
      Else
        YFunction._UpdateTimedReportCallbackList(sensor, False)
      End If
      Me._timedReportCallbackGyro = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackGyro Is Nothing)) Then
        Me._timedReportCallbackGyro(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function

    Public Overridable Function _loadQuaternion() As Integer
      Dim now_stamp As Integer = 0
      Dim age_ms As Integer = 0
      now_stamp = CType(((YAPI.GetTickCount()) And (&H7FFFFFFF)), Integer)
      age_ms = (((now_stamp - Me._qt_stamp)) And (&H7FFFFFFF))
      If ((age_ms >= 10) Or (Me._qt_stamp = 0)) Then
        If (Me.load(10) <> YAPI.SUCCESS) Then
          Return YAPI.DEVICE_NOT_FOUND
        End If
        If (Me._qt_stamp = 0) Then
          Me._qt_w = YQt.FindQt("" + Me._serial + ".qt1")
          Me._qt_x = YQt.FindQt("" + Me._serial + ".qt2")
          Me._qt_y = YQt.FindQt("" + Me._serial + ".qt3")
          Me._qt_z = YQt.FindQt("" + Me._serial + ".qt4")
        End If
        REM
        If (Me._qt_w.load(9) <> YAPI.SUCCESS) Then
          Return YAPI.DEVICE_NOT_FOUND
        End If
        If (Me._qt_x.load(9) <> YAPI.SUCCESS) Then
          Return YAPI.DEVICE_NOT_FOUND
        End If
        If (Me._qt_y.load(9) <> YAPI.SUCCESS) Then
          Return YAPI.DEVICE_NOT_FOUND
        End If
        If (Me._qt_z.load(9) <> YAPI.SUCCESS) Then
          Return YAPI.DEVICE_NOT_FOUND
        End If
        Me._w = Me._qt_w.get_currentValue()
        Me._x = Me._qt_x.get_currentValue()
        Me._y = Me._qt_y.get_currentValue()
        Me._z = Me._qt_z.get_currentValue()
        Me._qt_stamp = now_stamp
      End If
      Return YAPI.SUCCESS
    End Function

    Public Overridable Function _loadAngles() As Integer
      Dim sqw As Double = 0
      Dim sqx As Double = 0
      Dim sqy As Double = 0
      Dim sqz As Double = 0
      Dim norm As Double = 0
      Dim delta As Double = 0
      REM // may throw an exception
      If (Me._loadQuaternion() <> YAPI.SUCCESS) Then
        Return YAPI.DEVICE_NOT_FOUND
      End If
      If (Me._angles_stamp <> Me._qt_stamp) Then
        sqw = Me._w * Me._w
        sqx = Me._x * Me._x
        sqy = Me._y * Me._y
        sqz = Me._z * Me._z
        norm = sqx + sqy + sqz + sqw
        delta = Me._y * Me._w - Me._x * Me._z
        If (delta > 0.499 * norm) Then
          REM
          Me._pitch = 90.0
          Me._head  = Math.Round(2.0 * 1800.0/Math.PI * Math.Atan2(Me._x,Me._w)) / 10.0
        Else
          If (delta < -0.499 * norm) Then
            REM
            Me._pitch = -90.0
            Me._head  = Math.Round(-2.0 * 1800.0/Math.PI * Math.Atan2(Me._x,Me._w)) / 10.0
          Else
            Me._roll  = Math.Round(1800.0/Math.PI * Math.Atan2(2.0 * (Me._w * Me._x + Me._y * Me._z),sqw - sqx - sqy + sqz)) / 10.0
            Me._pitch = Math.Round(1800.0/Math.PI * Math.Asin(2.0 * delta / norm)) / 10.0
            Me._head  = Math.Round(1800.0/Math.PI * Math.Atan2(2.0 * (Me._x * Me._y + Me._z * Me._w),sqw + sqx - sqy - sqz)) / 10.0
          End If
        End If
        Me._angles_stamp = Me._qt_stamp
      End If
      Return YAPI.SUCCESS
    End Function

    '''*
    ''' <summary>
    '''   Returns the estimated roll angle, based on the integration of
    '''   gyroscopic measures combined with acceleration and
    '''   magnetic field measurements.
    ''' <para>
    '''   The axis corresponding to the roll angle can be mapped to any
    '''   of the device X, Y or Z physical directions using methods of
    '''   the class <c>YRefFrame</c>.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating-point number corresponding to roll angle
    '''   in degrees, between -180 and +180.
    ''' </returns>
    '''/
    Public Overridable Function get_roll() As Double
      REM // may throw an exception
      Me._loadAngles()
      Return Me._roll
    End Function

    '''*
    ''' <summary>
    '''   Returns the estimated pitch angle, based on the integration of
    '''   gyroscopic measures combined with acceleration and
    '''   magnetic field measurements.
    ''' <para>
    '''   The axis corresponding to the pitch angle can be mapped to any
    '''   of the device X, Y or Z physical directions using methods of
    '''   the class <c>YRefFrame</c>.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating-point number corresponding to pitch angle
    '''   in degrees, between -90 and +90.
    ''' </returns>
    '''/
    Public Overridable Function get_pitch() As Double
      REM // may throw an exception
      Me._loadAngles()
      Return Me._pitch
    End Function

    '''*
    ''' <summary>
    '''   Returns the estimated heading angle, based on the integration of
    '''   gyroscopic measures combined with acceleration and
    '''   magnetic field measurements.
    ''' <para>
    '''   The axis corresponding to the heading can be mapped to any
    '''   of the device X, Y or Z physical directions using methods of
    '''   the class <c>YRefFrame</c>.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating-point number corresponding to heading
    '''   in degrees, between 0 and 360.
    ''' </returns>
    '''/
    Public Overridable Function get_heading() As Double
      REM // may throw an exception
      Me._loadAngles()
      Return Me._head
    End Function

    '''*
    ''' <summary>
    '''   Returns the <c>w</c> component (real part) of the quaternion
    '''   describing the device estimated orientation, based on the
    '''   integration of gyroscopic measures combined with acceleration and
    '''   magnetic field measurements.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating-point number corresponding to the <c>w</c>
    '''   component of the quaternion.
    ''' </returns>
    '''/
    Public Overridable Function get_quaternionW() As Double
      REM // may throw an exception
      Me._loadQuaternion()
      Return Me._w
    End Function

    '''*
    ''' <summary>
    '''   Returns the <c>x</c> component of the quaternion
    '''   describing the device estimated orientation, based on the
    '''   integration of gyroscopic measures combined with acceleration and
    '''   magnetic field measurements.
    ''' <para>
    '''   The <c>x</c> component is
    '''   mostly correlated with rotations on the roll axis.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating-point number corresponding to the <c>x</c>
    '''   component of the quaternion.
    ''' </returns>
    '''/
    Public Overridable Function get_quaternionX() As Double
      REM // may throw an exception
      Me._loadQuaternion()
      Return Me._x
    End Function

    '''*
    ''' <summary>
    '''   Returns the <c>y</c> component of the quaternion
    '''   describing the device estimated orientation, based on the
    '''   integration of gyroscopic measures combined with acceleration and
    '''   magnetic field measurements.
    ''' <para>
    '''   The <c>y</c> component is
    '''   mostly correlated with rotations on the pitch axis.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating-point number corresponding to the <c>y</c>
    '''   component of the quaternion.
    ''' </returns>
    '''/
    Public Overridable Function get_quaternionY() As Double
      REM // may throw an exception
      Me._loadQuaternion()
      Return Me._y
    End Function

    '''*
    ''' <summary>
    '''   Returns the <c>x</c> component of the quaternion
    '''   describing the device estimated orientation, based on the
    '''   integration of gyroscopic measures combined with acceleration and
    '''   magnetic field measurements.
    ''' <para>
    '''   The <c>x</c> component is
    '''   mostly correlated with changes of heading.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating-point number corresponding to the <c>z</c>
    '''   component of the quaternion.
    ''' </returns>
    '''/
    Public Overridable Function get_quaternionZ() As Double
      REM // may throw an exception
      Me._loadQuaternion()
      Return Me._z
    End Function

    '''*
    ''' <summary>
    '''   Registers a callback function that will be invoked each time that the estimated
    '''   device orientation has changed.
    ''' <para>
    '''   The call frequency is typically around 95Hz during a move.
    '''   The callback is invoked only during the execution of <c>ySleep</c> or <c>yHandleEvents</c>.
    '''   This provides control over the time when the callback is triggered.
    '''   For good responsiveness, remember to call one of these two functions periodically.
    '''   To unregister a callback, pass a null pointer as argument.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="callback">
    '''   the callback function to invoke, or a null pointer.
    '''   The callback function should take five arguments:
    '''   the YGyro object of the turning device, and the floating
    '''   point values of the four components w, x, y and z
    '''   (as floating-point numbers).
    ''' @noreturn
    ''' </param>
    '''/
    Public Overridable Function registerQuaternionCallback(callback As YQuatCallback) As Integer
      Me._quatCallback = callback
      If (Not (callback Is Nothing)) Then
        REM
        If (Me._loadQuaternion() <> YAPI.SUCCESS) Then
          Return YAPI.DEVICE_NOT_FOUND
        End If
        Me._qt_w.set_userData(Me)
        Me._qt_x.set_userData(Me)
        Me._qt_y.set_userData(Me)
        Me._qt_z.set_userData(Me)
        Me._qt_w.registerValueCallback(AddressOf yInternalGyroCallback)
        Me._qt_x.registerValueCallback(AddressOf yInternalGyroCallback)
        Me._qt_y.registerValueCallback(AddressOf yInternalGyroCallback)
        Me._qt_z.registerValueCallback(AddressOf yInternalGyroCallback)
      Else
        If (Not (Not (Me._anglesCallback Is Nothing))) Then
          Me._qt_w.registerValueCallback(CType(Nothing, YQtValueCallback))
          Me._qt_x.registerValueCallback(CType(Nothing, YQtValueCallback))
          Me._qt_y.registerValueCallback(CType(Nothing, YQtValueCallback))
          Me._qt_z.registerValueCallback(CType(Nothing, YQtValueCallback))
        End If
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Registers a callback function that will be invoked each time that the estimated
    '''   device orientation has changed.
    ''' <para>
    '''   The call frequency is typically around 95Hz during a move.
    '''   The callback is invoked only during the execution of <c>ySleep</c> or <c>yHandleEvents</c>.
    '''   This provides control over the time when the callback is triggered.
    '''   For good responsiveness, remember to call one of these two functions periodically.
    '''   To unregister a callback, pass a null pointer as argument.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="callback">
    '''   the callback function to invoke, or a null pointer.
    '''   The callback function should take four arguments:
    '''   the YGyro object of the turning device, and the floating
    '''   point values of the three angles roll, pitch and heading
    '''   in degrees (as floating-point numbers).
    ''' @noreturn
    ''' </param>
    '''/
    Public Overridable Function registerAnglesCallback(callback As YAnglesCallback) As Integer
      Me._anglesCallback = callback
      If (Not (callback Is Nothing)) Then
        REM
        If (Me._loadQuaternion() <> YAPI.SUCCESS) Then
          Return YAPI.DEVICE_NOT_FOUND
        End If
        Me._qt_w.set_userData(Me)
        Me._qt_x.set_userData(Me)
        Me._qt_y.set_userData(Me)
        Me._qt_z.set_userData(Me)
        Me._qt_w.registerValueCallback(AddressOf yInternalGyroCallback)
        Me._qt_x.registerValueCallback(AddressOf yInternalGyroCallback)
        Me._qt_y.registerValueCallback(AddressOf yInternalGyroCallback)
        Me._qt_z.registerValueCallback(AddressOf yInternalGyroCallback)
      Else
        If (Not (Not (Me._quatCallback Is Nothing))) Then
          Me._qt_w.registerValueCallback(CType(Nothing, YQtValueCallback))
          Me._qt_x.registerValueCallback(CType(Nothing, YQtValueCallback))
          Me._qt_y.registerValueCallback(CType(Nothing, YQtValueCallback))
          Me._qt_z.registerValueCallback(CType(Nothing, YQtValueCallback))
        End If
      End If
      Return 0
    End Function

    Public Overridable Function _invokeGyroCallbacks(qtIndex As Integer, qtValue As Double) As Integer
      If (qtIndex - 1 = 0) Then
        Me._w = qtValue
      ElseIf (qtIndex - 1 = 1) Then
        Me._x = qtValue
      ElseIf (qtIndex - 1 = 2) Then
        Me._y = qtValue
      ElseIf (qtIndex - 1 = 3) Then
        Me._z = qtValue
      End If
      If (qtIndex < 4) Then
        Return 0
      End If
      Me._qt_stamp = CType(((YAPI.GetTickCount()) And (&H7FFFFFFF)), Integer)
      If (Not (Me._quatCallback Is Nothing)) Then
        Me._quatCallback(Me, Me._w, Me._x, Me._y, Me._z)
      End If
      If (Not (Me._anglesCallback Is Nothing)) Then
        REM
        Me._loadAngles()
        Me._anglesCallback(Me, Me._roll, Me._pitch, Me._head)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of gyroscopes started using <c>yFirstGyro()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YGyro</c> object, corresponding to
    '''   a gyroscope currently online, or a <c>null</c> pointer
    '''   if there are no more gyroscopes to enumerate.
    ''' </returns>
    '''/
    Public Function nextGyro() As YGyro
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YGyro.FindGyro(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of gyroscopes currently accessible.
    ''' <para>
    '''   Use the method <c>YGyro.nextGyro()</c> to iterate on
    '''   next gyroscopes.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YGyro</c> object, corresponding to
    '''   the first gyro currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstGyro() As YGyro
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Gyro", 0, p, size, neededsize, errmsg)
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
      Return YGyro.FindGyro(serial + "." + funcId)
    End Function

    REM --- (end of generated code: YGyro public methods declaration)

  End Class

  REM --- (generated code: Gyro functions)

  '''*
  ''' <summary>
  '''   Retrieves a gyroscope for a given identifier.
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
  '''   This function does not require that the gyroscope is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YGyro.isOnline()</c> to test if the gyroscope is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a gyroscope by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the gyroscope
  ''' </param>
  ''' <returns>
  '''   a <c>YGyro</c> object allowing you to drive the gyroscope.
  ''' </returns>
  '''/
  Public Function yFindGyro(ByVal func As String) As YGyro
    Return YGyro.FindGyro(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of gyroscopes currently accessible.
  ''' <para>
  '''   Use the method <c>YGyro.nextGyro()</c> to iterate on
  '''   next gyroscopes.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YGyro</c> object, corresponding to
  '''   the first gyro currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstGyro() As YGyro
    Return YGyro.FirstGyro()
  End Function


  REM --- (end of generated code: Gyro functions)

End Module

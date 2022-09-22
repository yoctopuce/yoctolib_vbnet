' ********************************************************************
'
'  $Id: yocto_magnetometer.vb 50689 2022-08-17 14:37:15Z mvuilleu $
'
'  Implements yFindMagnetometer(), the high-level API for Magnetometer functions
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

Module yocto_magnetometer

    REM --- (YMagnetometer return codes)
    REM --- (end of YMagnetometer return codes)
    REM --- (YMagnetometer dlldef)
    REM --- (end of YMagnetometer dlldef)
   REM --- (YMagnetometer yapiwrapper)
   REM --- (end of YMagnetometer yapiwrapper)
  REM --- (YMagnetometer globals)

  Public Const Y_BANDWIDTH_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_XVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_YVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_ZVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Delegate Sub YMagnetometerValueCallback(ByVal func As YMagnetometer, ByVal value As String)
  Public Delegate Sub YMagnetometerTimedReportCallback(ByVal func As YMagnetometer, ByVal measure As YMeasure)
  REM --- (end of YMagnetometer globals)

  REM --- (YMagnetometer class start)

  '''*
  ''' <summary>
  '''   The <c>YSensor</c> class is the parent class for all Yoctopuce sensor types.
  ''' <para>
  '''   It can be
  '''   used to read the current value and unit of any sensor, read the min/max
  '''   value, configure autonomous recording frequency and access recorded data.
  '''   It also provide a function to register a callback invoked each time the
  '''   observed value changes, or at a predefined interval. Using this class rather
  '''   than a specific subclass makes it possible to create generic applications
  '''   that work with any Yoctopuce sensor, even those that do not yet exist.
  '''   Note: The <c>YAnButton</c> class is the only analog input which does not inherit
  '''   from <c>YSensor</c>.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YMagnetometer
    Inherits YSensor
    REM --- (end of YMagnetometer class start)

    REM --- (YMagnetometer definitions)
    Public Const BANDWIDTH_INVALID As Integer = YAPI.INVALID_UINT
    Public Const XVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const YVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const ZVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
    REM --- (end of YMagnetometer definitions)

    REM --- (YMagnetometer attributes declaration)
    Protected _bandwidth As Integer
    Protected _xValue As Double
    Protected _yValue As Double
    Protected _zValue As Double
    Protected _valueCallbackMagnetometer As YMagnetometerValueCallback
    Protected _timedReportCallbackMagnetometer As YMagnetometerTimedReportCallback
    REM --- (end of YMagnetometer attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Magnetometer"
      REM --- (YMagnetometer attributes initialization)
      _bandwidth = BANDWIDTH_INVALID
      _xValue = XVALUE_INVALID
      _yValue = YVALUE_INVALID
      _zValue = ZVALUE_INVALID
      _valueCallbackMagnetometer = Nothing
      _timedReportCallbackMagnetometer = Nothing
      REM --- (end of YMagnetometer attributes initialization)
    End Sub

    REM --- (YMagnetometer private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("bandwidth") Then
        _bandwidth = CInt(json_val.getLong("bandwidth"))
      End If
      If json_val.has("xValue") Then
        _xValue = Math.Round(json_val.getDouble("xValue") / 65.536) / 1000.0
      End If
      If json_val.has("yValue") Then
        _yValue = Math.Round(json_val.getDouble("yValue") / 65.536) / 1000.0
      End If
      If json_val.has("zValue") Then
        _zValue = Math.Round(json_val.getDouble("zValue") / 65.536) / 1000.0
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YMagnetometer private methods declaration)

    REM --- (YMagnetometer public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the measure update frequency, measured in Hz.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the measure update frequency, measured in Hz
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMagnetometer.BANDWIDTH_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_bandwidth() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return BANDWIDTH_INVALID
        End If
      End If
      res = Me._bandwidth
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the measure update frequency, measured in Hz.
    ''' <para>
    '''   When the
    '''   frequency is lower, the device performs averaging.
    '''   Remember to call the <c>saveToFlash()</c>
    '''   method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the measure update frequency, measured in Hz
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
    Public Function set_bandwidth(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("bandwidth", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the X component of the magnetic field, as a floating point number.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the X component of the magnetic field, as a floating point number
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMagnetometer.XVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_xValue() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return XVALUE_INVALID
        End If
      End If
      res = Me._xValue
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the Y component of the magnetic field, as a floating point number.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the Y component of the magnetic field, as a floating point number
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMagnetometer.YVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_yValue() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return YVALUE_INVALID
        End If
      End If
      res = Me._yValue
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the Z component of the magnetic field, as a floating point number.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the Z component of the magnetic field, as a floating point number
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMagnetometer.ZVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_zValue() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return ZVALUE_INVALID
        End If
      End If
      res = Me._zValue
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Retrieves a magnetometer for a given identifier.
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
    '''   This function does not require that the magnetometer is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YMagnetometer.isOnline()</c> to test if the magnetometer is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a magnetometer by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the magnetometer, for instance
    '''   <c>Y3DMK002.magnetometer</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YMagnetometer</c> object allowing you to drive the magnetometer.
    ''' </returns>
    '''/
    Public Shared Function FindMagnetometer(func As String) As YMagnetometer
      Dim obj As YMagnetometer
      obj = CType(YFunction._FindFromCache("Magnetometer", func), YMagnetometer)
      If ((obj Is Nothing)) Then
        obj = New YMagnetometer(func)
        YFunction._AddToCache("Magnetometer", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YMagnetometerValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackMagnetometer = callback
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
      If (Not (Me._valueCallbackMagnetometer Is Nothing)) Then
        Me._valueCallbackMagnetometer(Me, value)
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
    Public Overloads Function registerTimedReportCallback(callback As YMagnetometerTimedReportCallback) As Integer
      Dim sensor As YSensor
      sensor = Me
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(sensor, True)
      Else
        YFunction._UpdateTimedReportCallbackList(sensor, False)
      End If
      Me._timedReportCallbackMagnetometer = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackMagnetometer Is Nothing)) Then
        Me._timedReportCallbackMagnetometer(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of magnetometers started using <c>yFirstMagnetometer()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned magnetometers order.
    '''   If you want to find a specific a magnetometer, use <c>Magnetometer.findMagnetometer()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YMagnetometer</c> object, corresponding to
    '''   a magnetometer currently online, or a <c>Nothing</c> pointer
    '''   if there are no more magnetometers to enumerate.
    ''' </returns>
    '''/
    Public Function nextMagnetometer() As YMagnetometer
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YMagnetometer.FindMagnetometer(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of magnetometers currently accessible.
    ''' <para>
    '''   Use the method <c>YMagnetometer.nextMagnetometer()</c> to iterate on
    '''   next magnetometers.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YMagnetometer</c> object, corresponding to
    '''   the first magnetometer currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstMagnetometer() As YMagnetometer
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Magnetometer", 0, p, size, neededsize, errmsg)
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
      Return YMagnetometer.FindMagnetometer(serial + "." + funcId)
    End Function

    REM --- (end of YMagnetometer public methods declaration)

  End Class

  REM --- (YMagnetometer functions)

  '''*
  ''' <summary>
  '''   Retrieves a magnetometer for a given identifier.
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
  '''   This function does not require that the magnetometer is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YMagnetometer.isOnline()</c> to test if the magnetometer is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a magnetometer by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the magnetometer, for instance
  '''   <c>Y3DMK002.magnetometer</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YMagnetometer</c> object allowing you to drive the magnetometer.
  ''' </returns>
  '''/
  Public Function yFindMagnetometer(ByVal func As String) As YMagnetometer
    Return YMagnetometer.FindMagnetometer(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of magnetometers currently accessible.
  ''' <para>
  '''   Use the method <c>YMagnetometer.nextMagnetometer()</c> to iterate on
  '''   next magnetometers.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YMagnetometer</c> object, corresponding to
  '''   the first magnetometer currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstMagnetometer() As YMagnetometer
    Return YMagnetometer.FirstMagnetometer()
  End Function


  REM --- (end of YMagnetometer functions)

End Module

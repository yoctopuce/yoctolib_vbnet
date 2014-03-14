'*********************************************************************
'*
'* $Id: yocto_magnetometer.vb 15259 2014-03-06 10:21:05Z seb $
'*
'* Implements yFindMagnetometer(), the high-level API for Magnetometer functions
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

Module yocto_magnetometer

    REM --- (YMagnetometer return codes)
    REM --- (end of YMagnetometer return codes)
  REM --- (YMagnetometer globals)

  Public Const Y_XVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_YVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_ZVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Delegate Sub YMagnetometerValueCallback(ByVal func As YMagnetometer, ByVal value As String)
  Public Delegate Sub YMagnetometerTimedReportCallback(ByVal func As YMagnetometer, ByVal measure As YMeasure)
  REM --- (end of YMagnetometer globals)

  REM --- (YMagnetometer class start)

  '''*
  ''' <summary>
  '''   The Yoctopuce application programming interface allows you to read an instant
  '''   measure of the sensor, as well as the minimal and maximal values observed.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YMagnetometer
    Inherits YSensor
    REM --- (end of YMagnetometer class start)

    REM --- (YMagnetometer definitions)
    Public Const XVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const YVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const ZVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
    REM --- (end of YMagnetometer definitions)

    REM --- (YMagnetometer attributes declaration)
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
      _xValue = XVALUE_INVALID
      _yValue = YVALUE_INVALID
      _zValue = ZVALUE_INVALID
      _valueCallbackMagnetometer = Nothing
      _timedReportCallbackMagnetometer = Nothing
      REM --- (end of YMagnetometer attributes initialization)
    End Sub

    REM --- (YMagnetometer private methods declaration)

    Protected Overrides Function _parseAttr(ByRef member As TJSONRECORD) As Integer
      If (member.name = "xValue") Then
        _xValue = member.ivalue / 65536.0
        Return 1
      End If
      If (member.name = "yValue") Then
        _yValue = member.ivalue / 65536.0
        Return 1
      End If
      If (member.name = "zValue") Then
        _zValue = member.ivalue / 65536.0
        Return 1
      End If
      Return MyBase._parseAttr(member)
    End Function

    REM --- (end of YMagnetometer private methods declaration)

    REM --- (YMagnetometer public methods declaration)
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
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the magnetometer
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
    Public Overloads Function registerValueCallback(callback As YMagnetometerValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackMagnetometer = callback
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
    Public Overloads Function registerTimedReportCallback(callback As YMagnetometerTimedReportCallback) As Integer
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(Me, True)
      Else
        YFunction._UpdateTimedReportCallbackList(Me, False)
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
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YMagnetometer</c> object, corresponding to
    '''   a magnetometer currently online, or a <c>null</c> pointer
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
    '''   the first magnetometer currently online, or a <c>null</c> pointer
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

  REM --- (Magnetometer functions)

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
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the magnetometer
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
  '''   the first magnetometer currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstMagnetometer() As YMagnetometer
    Return YMagnetometer.FirstMagnetometer()
  End Function


  REM --- (end of Magnetometer functions)

End Module

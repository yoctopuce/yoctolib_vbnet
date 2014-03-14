'*********************************************************************
'*
'* $Id: yocto_genericsensor.vb 15259 2014-03-06 10:21:05Z seb $
'*
'* Implements yFindGenericSensor(), the high-level API for GenericSensor functions
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

Module yocto_genericsensor

    REM --- (YGenericSensor return codes)
    REM --- (end of YGenericSensor return codes)
  REM --- (YGenericSensor globals)

  Public Const Y_SIGNALVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_SIGNALUNIT_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_SIGNALRANGE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_VALUERANGE_INVALID As String = YAPI.INVALID_STRING
  Public Delegate Sub YGenericSensorValueCallback(ByVal func As YGenericSensor, ByVal value As String)
  Public Delegate Sub YGenericSensorTimedReportCallback(ByVal func As YGenericSensor, ByVal measure As YMeasure)
  REM --- (end of YGenericSensor globals)

  REM --- (YGenericSensor class start)

  '''*
  ''' <summary>
  '''   The Yoctopuce application programming interface allows you to read an instant
  '''   measure of the sensor, as well as the minimal and maximal values observed.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YGenericSensor
    Inherits YSensor
    REM --- (end of YGenericSensor class start)

    REM --- (YGenericSensor definitions)
    Public Const SIGNALVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const SIGNALUNIT_INVALID As String = YAPI.INVALID_STRING
    Public Const SIGNALRANGE_INVALID As String = YAPI.INVALID_STRING
    Public Const VALUERANGE_INVALID As String = YAPI.INVALID_STRING
    REM --- (end of YGenericSensor definitions)

    REM --- (YGenericSensor attributes declaration)
    Protected _signalValue As Double
    Protected _signalUnit As String
    Protected _signalRange As String
    Protected _valueRange As String
    Protected _valueCallbackGenericSensor As YGenericSensorValueCallback
    Protected _timedReportCallbackGenericSensor As YGenericSensorTimedReportCallback
    REM --- (end of YGenericSensor attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "GenericSensor"
      REM --- (YGenericSensor attributes initialization)
      _signalValue = SIGNALVALUE_INVALID
      _signalUnit = SIGNALUNIT_INVALID
      _signalRange = SIGNALRANGE_INVALID
      _valueRange = VALUERANGE_INVALID
      _valueCallbackGenericSensor = Nothing
      _timedReportCallbackGenericSensor = Nothing
      REM --- (end of YGenericSensor attributes initialization)
    End Sub

    REM --- (YGenericSensor private methods declaration)

    Protected Overrides Function _parseAttr(ByRef member As TJSONRECORD) As Integer
      If (member.name = "signalValue") Then
        _signalValue = member.ivalue / 65536.0
        Return 1
      End If
      If (member.name = "signalUnit") Then
        _signalUnit = member.svalue
        Return 1
      End If
      If (member.name = "signalRange") Then
        _signalRange = member.svalue
        Return 1
      End If
      If (member.name = "valueRange") Then
        _valueRange = member.svalue
        Return 1
      End If
      Return MyBase._parseAttr(member)
    End Function

    REM --- (end of YGenericSensor private methods declaration)

    REM --- (YGenericSensor public methods declaration)

    '''*
    ''' <summary>
    '''   Changes the measuring unit for the measured value.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the measuring unit for the measured value
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
    '''   Returns the measured value of the electrical signal used by the sensor.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the measured value of the electrical signal used by the sensor
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_SIGNALVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_signalValue() As Double
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return SIGNALVALUE_INVALID
        End If
      End If
      Return Math.Round(Me._signalValue * 1000) / 1000
    End Function

    '''*
    ''' <summary>
    '''   Returns the measuring unit of the electrical signal used by the sensor.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the measuring unit of the electrical signal used by the sensor
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_SIGNALUNIT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_signalUnit() As String
      If (Me._cacheExpiration = 0) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return SIGNALUNIT_INVALID
        End If
      End If
      Return Me._signalUnit
    End Function

    '''*
    ''' <summary>
    '''   Returns the electric signal range used by the sensor.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the electric signal range used by the sensor
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_SIGNALRANGE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_signalRange() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return SIGNALRANGE_INVALID
        End If
      End If
      Return Me._signalRange
    End Function


    '''*
    ''' <summary>
    '''   Changes the electric signal range used by the sensor.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the electric signal range used by the sensor
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
    Public Function set_signalRange(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("signalRange", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the physical value range measured by the sensor.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the physical value range measured by the sensor
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_VALUERANGE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_valueRange() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return VALUERANGE_INVALID
        End If
      End If
      Return Me._valueRange
    End Function


    '''*
    ''' <summary>
    '''   Changes the physical value range measured by the sensor.
    ''' <para>
    '''   The range change may have a side effect
    '''   on the display resolution, as it may be adapted automatically.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the physical value range measured by the sensor
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
    Public Function set_valueRange(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("valueRange", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves a generic sensor for a given identifier.
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
    '''   This function does not require that the generic sensor is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YGenericSensor.isOnline()</c> to test if the generic sensor is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a generic sensor by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the generic sensor
    ''' </param>
    ''' <returns>
    '''   a <c>YGenericSensor</c> object allowing you to drive the generic sensor.
    ''' </returns>
    '''/
    Public Shared Function FindGenericSensor(func As String) As YGenericSensor
      Dim obj As YGenericSensor
      obj = CType(YFunction._FindFromCache("GenericSensor", func), YGenericSensor)
      If ((obj Is Nothing)) Then
        obj = New YGenericSensor(func)
        YFunction._AddToCache("GenericSensor", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YGenericSensorValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackGenericSensor = callback
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
      If (Not (Me._valueCallbackGenericSensor Is Nothing)) Then
        Me._valueCallbackGenericSensor(Me, value)
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
    Public Overloads Function registerTimedReportCallback(callback As YGenericSensorTimedReportCallback) As Integer
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(Me, True)
      Else
        YFunction._UpdateTimedReportCallbackList(Me, False)
      End If
      Me._timedReportCallbackGenericSensor = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackGenericSensor Is Nothing)) Then
        Me._timedReportCallbackGenericSensor(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of generic sensors started using <c>yFirstGenericSensor()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YGenericSensor</c> object, corresponding to
    '''   a generic sensor currently online, or a <c>null</c> pointer
    '''   if there are no more generic sensors to enumerate.
    ''' </returns>
    '''/
    Public Function nextGenericSensor() As YGenericSensor
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YGenericSensor.FindGenericSensor(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of generic sensors currently accessible.
    ''' <para>
    '''   Use the method <c>YGenericSensor.nextGenericSensor()</c> to iterate on
    '''   next generic sensors.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YGenericSensor</c> object, corresponding to
    '''   the first generic sensor currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstGenericSensor() As YGenericSensor
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("GenericSensor", 0, p, size, neededsize, errmsg)
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
      Return YGenericSensor.FindGenericSensor(serial + "." + funcId)
    End Function

    REM --- (end of YGenericSensor public methods declaration)

  End Class

  REM --- (GenericSensor functions)

  '''*
  ''' <summary>
  '''   Retrieves a generic sensor for a given identifier.
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
  '''   This function does not require that the generic sensor is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YGenericSensor.isOnline()</c> to test if the generic sensor is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a generic sensor by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the generic sensor
  ''' </param>
  ''' <returns>
  '''   a <c>YGenericSensor</c> object allowing you to drive the generic sensor.
  ''' </returns>
  '''/
  Public Function yFindGenericSensor(ByVal func As String) As YGenericSensor
    Return YGenericSensor.FindGenericSensor(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of generic sensors currently accessible.
  ''' <para>
  '''   Use the method <c>YGenericSensor.nextGenericSensor()</c> to iterate on
  '''   next generic sensors.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YGenericSensor</c> object, corresponding to
  '''   the first generic sensor currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstGenericSensor() As YGenericSensor
    Return YGenericSensor.FirstGenericSensor()
  End Function


  REM --- (end of GenericSensor functions)

End Module

' ********************************************************************
'
'  $Id: yocto_voltage.vb 37827 2019-10-25 13:07:48Z mvuilleu $
'
'  Implements yFindVoltage(), the high-level API for Voltage functions
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

Module yocto_voltage

    REM --- (YVoltage return codes)
    REM --- (end of YVoltage return codes)
    REM --- (YVoltage dlldef)
    REM --- (end of YVoltage dlldef)
   REM --- (YVoltage yapiwrapper)
   REM --- (end of YVoltage yapiwrapper)
  REM --- (YVoltage globals)

  Public Const Y_ENABLED_FALSE As Integer = 0
  Public Const Y_ENABLED_TRUE As Integer = 1
  Public Const Y_ENABLED_INVALID As Integer = -1
  Public Delegate Sub YVoltageValueCallback(ByVal func As YVoltage, ByVal value As String)
  Public Delegate Sub YVoltageTimedReportCallback(ByVal func As YVoltage, ByVal measure As YMeasure)
  REM --- (end of YVoltage globals)

  REM --- (YVoltage class start)

  '''*
  ''' <summary>
  '''   The YVoltage class allows you to read and configure Yoctopuce voltage
  '''   sensors, for instance using a Yocto-Watt, a Yocto-Volt or a Yocto-Motor-DC.
  ''' <para>
  '''   It inherits from YSensor class the core functions to read measurements,
  '''   to register callback functions, to access the autonomous datalogger.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YVoltage
    Inherits YSensor
    REM --- (end of YVoltage class start)

    REM --- (YVoltage definitions)
    Public Const ENABLED_FALSE As Integer = 0
    Public Const ENABLED_TRUE As Integer = 1
    Public Const ENABLED_INVALID As Integer = -1
    REM --- (end of YVoltage definitions)

    REM --- (YVoltage attributes declaration)
    Protected _enabled As Integer
    Protected _valueCallbackVoltage As YVoltageValueCallback
    Protected _timedReportCallbackVoltage As YVoltageTimedReportCallback
    REM --- (end of YVoltage attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Voltage"
      REM --- (YVoltage attributes initialization)
      _enabled = ENABLED_INVALID
      _valueCallbackVoltage = Nothing
      _timedReportCallbackVoltage = Nothing
      REM --- (end of YVoltage attributes initialization)
    End Sub

    REM --- (YVoltage private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("enabled") Then
        If (json_val.getInt("enabled") > 0) Then _enabled = 1 Else _enabled = 0
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YVoltage private methods declaration)

    REM --- (YVoltage public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the activation state of this input.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_ENABLED_FALSE</c> or <c>Y_ENABLED_TRUE</c>, according to the activation state of this input
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_ENABLED_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_enabled() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return ENABLED_INVALID
        End If
      End If
      res = Me._enabled
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the activation state of this voltage input.
    ''' <para>
    '''   When AC measurements are disabled,
    '''   the device will always assume a DC signal, and vice-versa. When both AC and DC measurements
    '''   are active, the device switches between AC and DC mode based on the relative amplitude
    '''   of variations compared to the average value.
    '''   Remember to call the <c>saveToFlash()</c>
    '''   method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>Y_ENABLED_FALSE</c> or <c>Y_ENABLED_TRUE</c>, according to the activation state of this voltage input
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
    Public Function set_enabled(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("enabled", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves a voltage sensor for a given identifier.
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
    '''   This function does not require that the voltage sensor is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YVoltage.isOnline()</c> to test if the voltage sensor is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a voltage sensor by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the voltage sensor, for instance
    '''   <c>YWATTMK1.voltage1</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YVoltage</c> object allowing you to drive the voltage sensor.
    ''' </returns>
    '''/
    Public Shared Function FindVoltage(func As String) As YVoltage
      Dim obj As YVoltage
      obj = CType(YFunction._FindFromCache("Voltage", func), YVoltage)
      If ((obj Is Nothing)) Then
        obj = New YVoltage(func)
        YFunction._AddToCache("Voltage", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YVoltageValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackVoltage = callback
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
      If (Not (Me._valueCallbackVoltage Is Nothing)) Then
        Me._valueCallbackVoltage(Me, value)
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
    Public Overloads Function registerTimedReportCallback(callback As YVoltageTimedReportCallback) As Integer
      Dim sensor As YSensor
      sensor = Me
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(sensor, True)
      Else
        YFunction._UpdateTimedReportCallbackList(sensor, False)
      End If
      Me._timedReportCallbackVoltage = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackVoltage Is Nothing)) Then
        Me._timedReportCallbackVoltage(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of voltage sensors started using <c>yFirstVoltage()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned voltage sensors order.
    '''   If you want to find a specific a voltage sensor, use <c>Voltage.findVoltage()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YVoltage</c> object, corresponding to
    '''   a voltage sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are no more voltage sensors to enumerate.
    ''' </returns>
    '''/
    Public Function nextVoltage() As YVoltage
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YVoltage.FindVoltage(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of voltage sensors currently accessible.
    ''' <para>
    '''   Use the method <c>YVoltage.nextVoltage()</c> to iterate on
    '''   next voltage sensors.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YVoltage</c> object, corresponding to
    '''   the first voltage sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstVoltage() As YVoltage
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Voltage", 0, p, size, neededsize, errmsg)
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
      Return YVoltage.FindVoltage(serial + "." + funcId)
    End Function

    REM --- (end of YVoltage public methods declaration)

  End Class

  REM --- (YVoltage functions)

  '''*
  ''' <summary>
  '''   Retrieves a voltage sensor for a given identifier.
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
  '''   This function does not require that the voltage sensor is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YVoltage.isOnline()</c> to test if the voltage sensor is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a voltage sensor by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the voltage sensor, for instance
  '''   <c>YWATTMK1.voltage1</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YVoltage</c> object allowing you to drive the voltage sensor.
  ''' </returns>
  '''/
  Public Function yFindVoltage(ByVal func As String) As YVoltage
    Return YVoltage.FindVoltage(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of voltage sensors currently accessible.
  ''' <para>
  '''   Use the method <c>YVoltage.nextVoltage()</c> to iterate on
  '''   next voltage sensors.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YVoltage</c> object, corresponding to
  '''   the first voltage sensor currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstVoltage() As YVoltage
    Return YVoltage.FirstVoltage()
  End Function


  REM --- (end of YVoltage functions)

End Module

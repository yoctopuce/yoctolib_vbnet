' ********************************************************************
'
'  $Id: yocto_power.vb 53431 2023-03-06 14:19:35Z seb $
'
'  Implements yFindPower(), the high-level API for Power functions
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

Module yocto_power

    REM --- (YPower return codes)
    REM --- (end of YPower return codes)
    REM --- (YPower dlldef)
    REM --- (end of YPower dlldef)
   REM --- (YPower yapiwrapper)
   REM --- (end of YPower yapiwrapper)
  REM --- (YPower globals)

  Public Const Y_POWERFACTOR_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_COSPHI_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_METER_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_DELIVEREDENERGYMETER_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_RECEIVEDENERGYMETER_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_METERTIMER_INVALID As Integer = YAPI.INVALID_UINT
  Public Delegate Sub YPowerValueCallback(ByVal func As YPower, ByVal value As String)
  Public Delegate Sub YPowerTimedReportCallback(ByVal func As YPower, ByVal measure As YMeasure)
  REM --- (end of YPower globals)

  REM --- (YPower class start)

  '''*
  ''' <summary>
  '''   The <c>YPower</c> class allows you to read and configure Yoctopuce electrical power sensors.
  ''' <para>
  '''   It inherits from <c>YSensor</c> class the core functions to read measurements,
  '''   to register callback functions, and to access the autonomous datalogger.
  '''   This class adds the ability to access the energy counter and the power factor.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YPower
    Inherits YSensor
    REM --- (end of YPower class start)

    REM --- (YPower definitions)
    Public Const POWERFACTOR_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const COSPHI_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const METER_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const DELIVEREDENERGYMETER_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const RECEIVEDENERGYMETER_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const METERTIMER_INVALID As Integer = YAPI.INVALID_UINT
    REM --- (end of YPower definitions)

    REM --- (YPower attributes declaration)
    Protected _powerFactor As Double
    Protected _cosPhi As Double
    Protected _meter As Double
    Protected _deliveredEnergyMeter As Double
    Protected _receivedEnergyMeter As Double
    Protected _meterTimer As Integer
    Protected _valueCallbackPower As YPowerValueCallback
    Protected _timedReportCallbackPower As YPowerTimedReportCallback
    REM --- (end of YPower attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Power"
      REM --- (YPower attributes initialization)
      _powerFactor = POWERFACTOR_INVALID
      _cosPhi = COSPHI_INVALID
      _meter = METER_INVALID
      _deliveredEnergyMeter = DELIVEREDENERGYMETER_INVALID
      _receivedEnergyMeter = RECEIVEDENERGYMETER_INVALID
      _meterTimer = METERTIMER_INVALID
      _valueCallbackPower = Nothing
      _timedReportCallbackPower = Nothing
      REM --- (end of YPower attributes initialization)
    End Sub

    REM --- (YPower private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("powerFactor") Then
        _powerFactor = Math.Round(json_val.getDouble("powerFactor") / 65.536) / 1000.0
      End If
      If json_val.has("cosPhi") Then
        _cosPhi = Math.Round(json_val.getDouble("cosPhi") / 65.536) / 1000.0
      End If
      If json_val.has("meter") Then
        _meter = Math.Round(json_val.getDouble("meter") / 65.536) / 1000.0
      End If
      If json_val.has("deliveredEnergyMeter") Then
        _deliveredEnergyMeter = Math.Round(json_val.getDouble("deliveredEnergyMeter") / 65.536) / 1000.0
      End If
      If json_val.has("receivedEnergyMeter") Then
        _receivedEnergyMeter = Math.Round(json_val.getDouble("receivedEnergyMeter") / 65.536) / 1000.0
      End If
      If json_val.has("meterTimer") Then
        _meterTimer = CInt(json_val.getLong("meterTimer"))
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YPower private methods declaration)

    REM --- (YPower public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the power factor (PF), i.e.
    ''' <para>
    '''   ratio between the active power consumed (in W)
    '''   and the apparent power provided (VA).
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the power factor (PF), i.e
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YPower.POWERFACTOR_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_powerFactor() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return POWERFACTOR_INVALID
        End If
      End If
      res = Me._powerFactor
      If (res = POWERFACTOR_INVALID) Then
        res = Me._cosPhi
      End If
      res = Math.Round(res * 1000) / 1000
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the Displacement Power factor (DPF), i.e.
    ''' <para>
    '''   cosine of the phase shift between
    '''   the voltage and current fundamentals.
    '''   On the Yocto-Watt (V1), the value returned by this method correponds to the
    '''   power factor as this device is cannot estimate the true DPF.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the Displacement Power factor (DPF), i.e
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YPower.COSPHI_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_cosPhi() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return COSPHI_INVALID
        End If
      End If
      res = Me._cosPhi
      Return res
    End Function


    Public Function set_meter(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("meter", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the energy counter, maintained by the wattmeter by integrating the
    '''   power consumption over time.
    ''' <para>
    '''   This is the sum of forward and backwad energy transfers,
    '''   if you are insterested in only one direction, use  get_receivedEnergyMeter() or
    '''   get_deliveredEnergyMeter(). Note that this counter is reset at each start of the device.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the energy counter, maintained by the wattmeter by integrating the
    '''   power consumption over time
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YPower.METER_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_meter() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return METER_INVALID
        End If
      End If
      res = Me._meter
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the energy counter, maintained by the wattmeter by integrating the power consumption over time,
    '''   but only when positive.
    ''' <para>
    '''   Note that this counter is reset at each start of the device.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the energy counter, maintained by the wattmeter by
    '''   integrating the power consumption over time,
    '''   but only when positive
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YPower.DELIVEREDENERGYMETER_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_deliveredEnergyMeter() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return DELIVEREDENERGYMETER_INVALID
        End If
      End If
      res = Me._deliveredEnergyMeter
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the energy counter, maintained by the wattmeter by integrating the power consumption over time,
    '''   but only when negative.
    ''' <para>
    '''   Note that this counter is reset at each start of the device.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the energy counter, maintained by the wattmeter by
    '''   integrating the power consumption over time,
    '''   but only when negative
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YPower.RECEIVEDENERGYMETER_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_receivedEnergyMeter() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return RECEIVEDENERGYMETER_INVALID
        End If
      End If
      res = Me._receivedEnergyMeter
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the elapsed time since last energy counter reset, in seconds.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the elapsed time since last energy counter reset, in seconds
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YPower.METERTIMER_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_meterTimer() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return METERTIMER_INVALID
        End If
      End If
      res = Me._meterTimer
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Retrieves a electrical power sensor for a given identifier.
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
    '''   This function does not require that the electrical power sensor is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YPower.isOnline()</c> to test if the electrical power sensor is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a electrical power sensor by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the electrical power sensor, for instance
    '''   <c>YWATTMK1.power</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YPower</c> object allowing you to drive the electrical power sensor.
    ''' </returns>
    '''/
    Public Shared Function FindPower(func As String) As YPower
      Dim obj As YPower
      obj = CType(YFunction._FindFromCache("Power", func), YPower)
      If ((obj Is Nothing)) Then
        obj = New YPower(func)
        YFunction._AddToCache("Power", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YPowerValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackPower = callback
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
      If (Not (Me._valueCallbackPower Is Nothing)) Then
        Me._valueCallbackPower(Me, value)
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
    Public Overloads Function registerTimedReportCallback(callback As YPowerTimedReportCallback) As Integer
      Dim sensor As YSensor
      sensor = Me
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(sensor, True)
      Else
        YFunction._UpdateTimedReportCallbackList(sensor, False)
      End If
      Me._timedReportCallbackPower = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackPower Is Nothing)) Then
        Me._timedReportCallbackPower(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Resets the energy counters.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function reset() As Integer
      Return Me.set_meter(0)
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of electrical power sensors started using <c>yFirstPower()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned electrical power sensors order.
    '''   If you want to find a specific a electrical power sensor, use <c>Power.findPower()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YPower</c> object, corresponding to
    '''   a electrical power sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are no more electrical power sensors to enumerate.
    ''' </returns>
    '''/
    Public Function nextPower() As YPower
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YPower.FindPower(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of electrical power sensors currently accessible.
    ''' <para>
    '''   Use the method <c>YPower.nextPower()</c> to iterate on
    '''   next electrical power sensors.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YPower</c> object, corresponding to
    '''   the first electrical power sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstPower() As YPower
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Power", 0, p, size, neededsize, errmsg)
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
      Return YPower.FindPower(serial + "." + funcId)
    End Function

    REM --- (end of YPower public methods declaration)

  End Class

  REM --- (YPower functions)

  '''*
  ''' <summary>
  '''   Retrieves a electrical power sensor for a given identifier.
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
  '''   This function does not require that the electrical power sensor is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YPower.isOnline()</c> to test if the electrical power sensor is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a electrical power sensor by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the electrical power sensor, for instance
  '''   <c>YWATTMK1.power</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YPower</c> object allowing you to drive the electrical power sensor.
  ''' </returns>
  '''/
  Public Function yFindPower(ByVal func As String) As YPower
    Return YPower.FindPower(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of electrical power sensors currently accessible.
  ''' <para>
  '''   Use the method <c>YPower.nextPower()</c> to iterate on
  '''   next electrical power sensors.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YPower</c> object, corresponding to
  '''   the first electrical power sensor currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstPower() As YPower
    Return YPower.FirstPower()
  End Function


  REM --- (end of YPower functions)

End Module

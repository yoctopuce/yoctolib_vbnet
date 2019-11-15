' ********************************************************************
'
'  $Id: yocto_humidity.vb 37827 2019-10-25 13:07:48Z mvuilleu $
'
'  Implements yFindHumidity(), the high-level API for Humidity functions
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

Module yocto_humidity

    REM --- (YHumidity return codes)
    REM --- (end of YHumidity return codes)
    REM --- (YHumidity dlldef)
    REM --- (end of YHumidity dlldef)
   REM --- (YHumidity yapiwrapper)
   REM --- (end of YHumidity yapiwrapper)
  REM --- (YHumidity globals)

  Public Const Y_RELHUM_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_ABSHUM_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Delegate Sub YHumidityValueCallback(ByVal func As YHumidity, ByVal value As String)
  Public Delegate Sub YHumidityTimedReportCallback(ByVal func As YHumidity, ByVal measure As YMeasure)
  REM --- (end of YHumidity globals)

  REM --- (YHumidity class start)

  '''*
  ''' <summary>
  '''   The YHumidity class allows you to read and configure Yoctopuce humidity
  '''   sensors, for instance using a Yocto-Meteo-V2, a Yocto-VOC-V3 or a Yocto-CO2-V2.
  ''' <para>
  '''   It inherits from YSensor class the core functions to read measurements,
  '''   to register callback functions, to access the autonomous datalogger.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YHumidity
    Inherits YSensor
    REM --- (end of YHumidity class start)

    REM --- (YHumidity definitions)
    Public Const RELHUM_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const ABSHUM_INVALID As Double = YAPI.INVALID_DOUBLE
    REM --- (end of YHumidity definitions)

    REM --- (YHumidity attributes declaration)
    Protected _relHum As Double
    Protected _absHum As Double
    Protected _valueCallbackHumidity As YHumidityValueCallback
    Protected _timedReportCallbackHumidity As YHumidityTimedReportCallback
    REM --- (end of YHumidity attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Humidity"
      REM --- (YHumidity attributes initialization)
      _relHum = RELHUM_INVALID
      _absHum = ABSHUM_INVALID
      _valueCallbackHumidity = Nothing
      _timedReportCallbackHumidity = Nothing
      REM --- (end of YHumidity attributes initialization)
    End Sub

    REM --- (YHumidity private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("relHum") Then
        _relHum = Math.Round(json_val.getDouble("relHum") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("absHum") Then
        _absHum = Math.Round(json_val.getDouble("absHum") * 1000.0 / 65536.0) / 1000.0
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YHumidity private methods declaration)

    REM --- (YHumidity public methods declaration)

    '''*
    ''' <summary>
    '''   Changes the primary unit for measuring humidity.
    ''' <para>
    '''   That unit is a string.
    '''   If that strings starts with the letter 'g', the primary measured value is the absolute
    '''   humidity, in g/m3. Otherwise, the primary measured value will be the relative humidity
    '''   (RH), in per cents.
    ''' </para>
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the modification
    '''   must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the primary unit for measuring humidity
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
    '''   Returns the current relative humidity, in per cents.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the current relative humidity, in per cents
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_RELHUM_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_relHum() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return RELHUM_INVALID
        End If
      End If
      res = Me._relHum
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the current absolute humidity, in grams per cubic meter of air.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the current absolute humidity, in grams per cubic meter of air
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_ABSHUM_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_absHum() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return ABSHUM_INVALID
        End If
      End If
      res = Me._absHum
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Retrieves a humidity sensor for a given identifier.
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
    '''   This function does not require that the humidity sensor is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YHumidity.isOnline()</c> to test if the humidity sensor is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a humidity sensor by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the humidity sensor, for instance
    '''   <c>METEOMK2.humidity</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YHumidity</c> object allowing you to drive the humidity sensor.
    ''' </returns>
    '''/
    Public Shared Function FindHumidity(func As String) As YHumidity
      Dim obj As YHumidity
      obj = CType(YFunction._FindFromCache("Humidity", func), YHumidity)
      If ((obj Is Nothing)) Then
        obj = New YHumidity(func)
        YFunction._AddToCache("Humidity", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YHumidityValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackHumidity = callback
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
      If (Not (Me._valueCallbackHumidity Is Nothing)) Then
        Me._valueCallbackHumidity(Me, value)
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
    Public Overloads Function registerTimedReportCallback(callback As YHumidityTimedReportCallback) As Integer
      Dim sensor As YSensor
      sensor = Me
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(sensor, True)
      Else
        YFunction._UpdateTimedReportCallbackList(sensor, False)
      End If
      Me._timedReportCallbackHumidity = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackHumidity Is Nothing)) Then
        Me._timedReportCallbackHumidity(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of humidity sensors started using <c>yFirstHumidity()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned humidity sensors order.
    '''   If you want to find a specific a humidity sensor, use <c>Humidity.findHumidity()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YHumidity</c> object, corresponding to
    '''   a humidity sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are no more humidity sensors to enumerate.
    ''' </returns>
    '''/
    Public Function nextHumidity() As YHumidity
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YHumidity.FindHumidity(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of humidity sensors currently accessible.
    ''' <para>
    '''   Use the method <c>YHumidity.nextHumidity()</c> to iterate on
    '''   next humidity sensors.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YHumidity</c> object, corresponding to
    '''   the first humidity sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstHumidity() As YHumidity
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Humidity", 0, p, size, neededsize, errmsg)
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
      Return YHumidity.FindHumidity(serial + "." + funcId)
    End Function

    REM --- (end of YHumidity public methods declaration)

  End Class

  REM --- (YHumidity functions)

  '''*
  ''' <summary>
  '''   Retrieves a humidity sensor for a given identifier.
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
  '''   This function does not require that the humidity sensor is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YHumidity.isOnline()</c> to test if the humidity sensor is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a humidity sensor by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the humidity sensor, for instance
  '''   <c>METEOMK2.humidity</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YHumidity</c> object allowing you to drive the humidity sensor.
  ''' </returns>
  '''/
  Public Function yFindHumidity(ByVal func As String) As YHumidity
    Return YHumidity.FindHumidity(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of humidity sensors currently accessible.
  ''' <para>
  '''   Use the method <c>YHumidity.nextHumidity()</c> to iterate on
  '''   next humidity sensors.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YHumidity</c> object, corresponding to
  '''   the first humidity sensor currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstHumidity() As YHumidity
    Return YHumidity.FirstHumidity()
  End Function


  REM --- (end of YHumidity functions)

End Module

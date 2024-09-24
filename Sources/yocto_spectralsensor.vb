' ********************************************************************
'
'  $Id: svn_id $
'
'  Implements yFindSpectralSensor(), the high-level API for SpectralSensor functions
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

Module yocto_spectralsensor

    REM --- (YSpectralSensor return codes)
    REM --- (end of YSpectralSensor return codes)
    REM --- (YSpectralSensor dlldef)
    REM --- (end of YSpectralSensor dlldef)
   REM --- (YSpectralSensor yapiwrapper)
   REM --- (end of YSpectralSensor yapiwrapper)
  REM --- (YSpectralSensor globals)

  Public Const Y_LEDCURRENT_INVALID As Integer = YAPI.INVALID_INT
  Public Const Y_RESOLUTION_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_INTEGRATIONTIME_INVALID As Integer = YAPI.INVALID_INT
  Public Const Y_GAIN_INVALID As Integer = YAPI.INVALID_INT
  Public Const Y_SATURATION_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_LEDCURRENTATPOWERON_INVALID As Integer = YAPI.INVALID_INT
  Public Const Y_INTEGRATIONTIMEATPOWERON_INVALID As Integer = YAPI.INVALID_INT
  Public Const Y_GAINATPOWERON_INVALID As Integer = YAPI.INVALID_INT
  Public Delegate Sub YSpectralSensorValueCallback(ByVal func As YSpectralSensor, ByVal value As String)
  Public Delegate Sub YSpectralSensorTimedReportCallback(ByVal func As YSpectralSensor, ByVal measure As YMeasure)
  REM --- (end of YSpectralSensor globals)

  REM --- (YSpectralSensor class start)

  '''*
  ''' <summary>
  '''   The <c>YSpectralSensor</c> class allows you to read and configure Yoctopuce spectral sensors.
  ''' <para>
  '''   It inherits from <c>YSensor</c> class the core functions to read measurements,
  '''   to register callback functions, and to access the autonomous datalogger.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YSpectralSensor
    Inherits YFunction
    REM --- (end of YSpectralSensor class start)

    REM --- (YSpectralSensor definitions)
    Public Const LEDCURRENT_INVALID As Integer = YAPI.INVALID_INT
    Public Const RESOLUTION_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const INTEGRATIONTIME_INVALID As Integer = YAPI.INVALID_INT
    Public Const GAIN_INVALID As Integer = YAPI.INVALID_INT
    Public Const SATURATION_INVALID As Integer = YAPI.INVALID_UINT
    Public Const LEDCURRENTATPOWERON_INVALID As Integer = YAPI.INVALID_INT
    Public Const INTEGRATIONTIMEATPOWERON_INVALID As Integer = YAPI.INVALID_INT
    Public Const GAINATPOWERON_INVALID As Integer = YAPI.INVALID_INT
    REM --- (end of YSpectralSensor definitions)

    REM --- (YSpectralSensor attributes declaration)
    Protected _ledCurrent As Integer
    Protected _resolution As Double
    Protected _integrationTime As Integer
    Protected _gain As Integer
    Protected _saturation As Integer
    Protected _ledCurrentAtPowerOn As Integer
    Protected _integrationTimeAtPowerOn As Integer
    Protected _gainAtPowerOn As Integer
    Protected _valueCallbackSpectralSensor As YSpectralSensorValueCallback
    REM --- (end of YSpectralSensor attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "SpectralSensor"
      REM --- (YSpectralSensor attributes initialization)
      _ledCurrent = LEDCURRENT_INVALID
      _resolution = RESOLUTION_INVALID
      _integrationTime = INTEGRATIONTIME_INVALID
      _gain = GAIN_INVALID
      _saturation = SATURATION_INVALID
      _ledCurrentAtPowerOn = LEDCURRENTATPOWERON_INVALID
      _integrationTimeAtPowerOn = INTEGRATIONTIMEATPOWERON_INVALID
      _gainAtPowerOn = GAINATPOWERON_INVALID
      _valueCallbackSpectralSensor = Nothing
      REM --- (end of YSpectralSensor attributes initialization)
    End Sub

    REM --- (YSpectralSensor private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("ledCurrent") Then
        _ledCurrent = CInt(json_val.getLong("ledCurrent"))
      End If
      If json_val.has("resolution") Then
        _resolution = Math.Round(json_val.getDouble("resolution") / 65.536) / 1000.0
      End If
      If json_val.has("integrationTime") Then
        _integrationTime = CInt(json_val.getLong("integrationTime"))
      End If
      If json_val.has("gain") Then
        _gain = CInt(json_val.getLong("gain"))
      End If
      If json_val.has("saturation") Then
        _saturation = CInt(json_val.getLong("saturation"))
      End If
      If json_val.has("ledCurrentAtPowerOn") Then
        _ledCurrentAtPowerOn = CInt(json_val.getLong("ledCurrentAtPowerOn"))
      End If
      If json_val.has("integrationTimeAtPowerOn") Then
        _integrationTimeAtPowerOn = CInt(json_val.getLong("integrationTimeAtPowerOn"))
      End If
      If json_val.has("gainAtPowerOn") Then
        _gainAtPowerOn = CInt(json_val.getLong("gainAtPowerOn"))
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YSpectralSensor private methods declaration)

    REM --- (YSpectralSensor public methods declaration)
    '''*
    ''' <summary>
    ''' </summary>
    ''' <returns>
    '''   an integer
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSpectralSensor.LEDCURRENT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_ledCurrent() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return LEDCURRENT_INVALID
        End If
      End If
      res = Me._ledCurrent
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the luminosity of the module leds.
    ''' <para>
    '''   The parameter is a
    '''   value between 0 and 100.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the luminosity of the module leds
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
    Public Function set_ledCurrent(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("ledCurrent", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Changes the resolution of the measured physical values.
    ''' <para>
    '''   The resolution corresponds to the numerical precision
    '''   when displaying value. It does not change the precision of the measure itself.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the resolution of the measured physical values
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
    Public Function set_resolution(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("resolution", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the resolution of the measured values.
    ''' <para>
    '''   The resolution corresponds to the numerical precision
    '''   of the measures, which is not always the same as the actual precision of the sensor.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the resolution of the measured values
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSpectralSensor.RESOLUTION_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_resolution() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return RESOLUTION_INVALID
        End If
      End If
      res = Me._resolution
      Return res
    End Function

    '''*
    ''' <summary>
    ''' </summary>
    ''' <returns>
    '''   an integer
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSpectralSensor.INTEGRATIONTIME_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_integrationTime() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return INTEGRATIONTIME_INVALID
        End If
      End If
      res = Me._integrationTime
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Change the integration time for a measure.
    ''' <para>
    '''   The parameter is a
    '''   value between 0 and 100.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer
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
    Public Function set_integrationTime(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("integrationTime", rest_val)
    End Function
    '''*
    ''' <summary>
    ''' </summary>
    ''' <returns>
    '''   an integer
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSpectralSensor.GAIN_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_gain() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return GAIN_INVALID
        End If
      End If
      res = Me._gain
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer
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
    Public Function set_gain(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("gain", rest_val)
    End Function
    '''*
    ''' <summary>
    ''' </summary>
    ''' <returns>
    '''   an integer
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSpectralSensor.SATURATION_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_saturation() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return SATURATION_INVALID
        End If
      End If
      res = Me._saturation
      Return res
    End Function

    '''*
    ''' <summary>
    ''' </summary>
    ''' <returns>
    '''   an integer
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSpectralSensor.LEDCURRENTATPOWERON_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_ledCurrentAtPowerOn() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return LEDCURRENTATPOWERON_INVALID
        End If
      End If
      res = Me._ledCurrentAtPowerOn
      Return res
    End Function


    '''*
    ''' <summary>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer
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
    Public Function set_ledCurrentAtPowerOn(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("ledCurrentAtPowerOn", rest_val)
    End Function
    '''*
    ''' <summary>
    ''' </summary>
    ''' <returns>
    '''   an integer
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSpectralSensor.INTEGRATIONTIMEATPOWERON_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_integrationTimeAtPowerOn() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return INTEGRATIONTIMEATPOWERON_INVALID
        End If
      End If
      res = Me._integrationTimeAtPowerOn
      Return res
    End Function


    '''*
    ''' <summary>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer
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
    Public Function set_integrationTimeAtPowerOn(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("integrationTimeAtPowerOn", rest_val)
    End Function
    '''*
    ''' <summary>
    ''' </summary>
    ''' <returns>
    '''   an integer
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSpectralSensor.GAINATPOWERON_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_gainAtPowerOn() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return GAINATPOWERON_INVALID
        End If
      End If
      res = Me._gainAtPowerOn
      Return res
    End Function


    '''*
    ''' <summary>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer
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
    Public Function set_gainAtPowerOn(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("gainAtPowerOn", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves a spectral sensor for a given identifier.
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
    '''   This function does not require that the spectral sensor is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YSpectralSensor.isOnline()</c> to test if the spectral sensor is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a spectral sensor by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the spectral sensor, for instance
    '''   <c>MyDevice.spectralSensor</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YSpectralSensor</c> object allowing you to drive the spectral sensor.
    ''' </returns>
    '''/
    Public Shared Function FindSpectralSensor(func As String) As YSpectralSensor
      Dim obj As YSpectralSensor
      obj = CType(YFunction._FindFromCache("SpectralSensor", func), YSpectralSensor)
      If ((obj Is Nothing)) Then
        obj = New YSpectralSensor(func)
        YFunction._AddToCache("SpectralSensor", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YSpectralSensorValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackSpectralSensor = callback
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
      If (Not (Me._valueCallbackSpectralSensor Is Nothing)) Then
        Me._valueCallbackSpectralSensor(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of spectral sensors started using <c>yFirstSpectralSensor()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned spectral sensors order.
    '''   If you want to find a specific a spectral sensor, use <c>SpectralSensor.findSpectralSensor()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YSpectralSensor</c> object, corresponding to
    '''   a spectral sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are no more spectral sensors to enumerate.
    ''' </returns>
    '''/
    Public Function nextSpectralSensor() As YSpectralSensor
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YSpectralSensor.FindSpectralSensor(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of spectral sensors currently accessible.
    ''' <para>
    '''   Use the method <c>YSpectralSensor.nextSpectralSensor()</c> to iterate on
    '''   next spectral sensors.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YSpectralSensor</c> object, corresponding to
    '''   the first spectral sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstSpectralSensor() As YSpectralSensor
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("SpectralSensor", 0, p, size, neededsize, errmsg)
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
      Return YSpectralSensor.FindSpectralSensor(serial + "." + funcId)
    End Function

    REM --- (end of YSpectralSensor public methods declaration)

  End Class

  REM --- (YSpectralSensor functions)

  '''*
  ''' <summary>
  '''   Retrieves a spectral sensor for a given identifier.
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
  '''   This function does not require that the spectral sensor is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YSpectralSensor.isOnline()</c> to test if the spectral sensor is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a spectral sensor by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the spectral sensor, for instance
  '''   <c>MyDevice.spectralSensor</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YSpectralSensor</c> object allowing you to drive the spectral sensor.
  ''' </returns>
  '''/
  Public Function yFindSpectralSensor(ByVal func As String) As YSpectralSensor
    Return YSpectralSensor.FindSpectralSensor(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of spectral sensors currently accessible.
  ''' <para>
  '''   Use the method <c>YSpectralSensor.nextSpectralSensor()</c> to iterate on
  '''   next spectral sensors.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YSpectralSensor</c> object, corresponding to
  '''   the first spectral sensor currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstSpectralSensor() As YSpectralSensor
    Return YSpectralSensor.FirstSpectralSensor()
  End Function


  REM --- (end of YSpectralSensor functions)

End Module

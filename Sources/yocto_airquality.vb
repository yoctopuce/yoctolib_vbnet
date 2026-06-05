' ********************************************************************
'
'  $Id: svn_id $
'
'  Implements yFindAirQuality(), the high-level API for AirQuality functions
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

Module yocto_airquality

    REM --- (YAirQuality return codes)
    REM --- (end of YAirQuality return codes)
    REM --- (YAirQuality dlldef)
    REM --- (end of YAirQuality dlldef)
   REM --- (YAirQuality yapiwrapper)
   REM --- (end of YAirQuality yapiwrapper)
  REM --- (YAirQuality globals)

  Public Const Y_UBAINDEX_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_RELATIVEINDEX_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_AQIMODE_RELATIVE As Integer = 0
  Public Const Y_AQIMODE_UBA As Integer = 1
  Public Const Y_AQIMODE_INVALID As Integer = -1
  Public Delegate Sub YAirQualityValueCallback(ByVal func As YAirQuality, ByVal value As String)
  Public Delegate Sub YAirQualityTimedReportCallback(ByVal func As YAirQuality, ByVal measure As YMeasure)
  REM --- (end of YAirQuality globals)

  REM --- (YAirQuality class start)

  '''*
  ''' <summary>
  '''   The <c>YAirQuality</c> class allows you to read and configure Yoctopuce air quality sensors.
  ''' <para>
  '''   It inherits from <c>YSensor</c> class the core functions to read measurements,
  '''   to register callback functions, and to access the autonomous datalogger.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YAirQuality
    Inherits YSensor
    REM --- (end of YAirQuality class start)

    REM --- (YAirQuality definitions)
    Public Const UBAINDEX_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const RELATIVEINDEX_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const AQIMODE_RELATIVE As Integer = 0
    Public Const AQIMODE_UBA As Integer = 1
    Public Const AQIMODE_INVALID As Integer = -1
    REM --- (end of YAirQuality definitions)

    REM --- (YAirQuality attributes declaration)
    Protected _ubaIndex As Double
    Protected _relativeIndex As Double
    Protected _aqiMode As Integer
    Protected _valueCallbackAirQuality As YAirQualityValueCallback
    Protected _timedReportCallbackAirQuality As YAirQualityTimedReportCallback
    REM --- (end of YAirQuality attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "AirQuality"
      REM --- (YAirQuality attributes initialization)
      _ubaIndex = UBAINDEX_INVALID
      _relativeIndex = RELATIVEINDEX_INVALID
      _aqiMode = AQIMODE_INVALID
      _valueCallbackAirQuality = Nothing
      _timedReportCallbackAirQuality = Nothing
      REM --- (end of YAirQuality attributes initialization)
    End Sub

    REM --- (YAirQuality private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("ubaIndex") Then
        _ubaIndex = Math.Round(json_val.getDouble("ubaIndex") / 65.536) / 1000.0
      End If
      If json_val.has("relativeIndex") Then
        _relativeIndex = Math.Round(json_val.getDouble("relativeIndex") / 65.536) / 1000.0
      End If
      If json_val.has("aqiMode") Then
        _aqiMode = CInt(json_val.getLong("aqiMode"))
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YAirQuality private methods declaration)

    REM --- (YAirQuality public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the current air quality index, according to UBA (from 1 to 5).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the current air quality index, according to UBA (from 1 to 5)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YAirQuality.UBAINDEX_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_ubaIndex() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return UBAINDEX_INVALID
        End If
      End If
      res = Me._ubaIndex
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the relative air quality index, according to ScioSense (from 0 to 500).
    ''' <para>
    '''   A value below 100 indicates better-than-average air quality compared to the past 24 hours,
    '''   while a value above 100 indicates poorer-than-average air quality compared to the past 24 hours.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the relative air quality index, according to ScioSense (from 0 to 500)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YAirQuality.RELATIVEINDEX_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_relativeIndex() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return RELATIVEINDEX_INVALID
        End If
      End If
      res = Me._relativeIndex
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the type of index reported by the get_currentValue function and callbacks (UBA index or relative index).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>YAirQuality.AQIMODE_RELATIVE</c> or <c>YAirQuality.AQIMODE_UBA</c>, according to the type
    '''   of index reported by the get_currentValue function and callbacks (UBA index or relative index)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YAirQuality.AQIMODE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_aqiMode() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return AQIMODE_INVALID
        End If
      End If
      res = Me._aqiMode
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the the type of index reported by the get_currentValue function and callbacks (UBA index or relative index).
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>YAirQuality.AQIMODE_RELATIVE</c> or <c>YAirQuality.AQIMODE_UBA</c>, according to the the
    '''   type of index reported by the get_currentValue function and callbacks (UBA index or relative index)
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
    Public Function set_aqiMode(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("aqiMode", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves a air quality sensor for a given identifier.
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
    '''   This function does not require that the air quality sensor is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YAirQuality.isOnline()</c> to test if the air quality sensor is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a air quality sensor by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the air quality sensor, for instance
    '''   <c>MyDevice.airQuality</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YAirQuality</c> object allowing you to drive the air quality sensor.
    ''' </returns>
    '''/
    Public Shared Function FindAirQuality(func As String) As YAirQuality
      Dim obj As YAirQuality
      obj = CType(YFunction._FindFromCache("AirQuality", func), YAirQuality)
      If ((obj Is Nothing)) Then
        obj = New YAirQuality(func)
        YFunction._AddToCache("AirQuality", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YAirQualityValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackAirQuality = callback
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
      If (Not (Me._valueCallbackAirQuality Is Nothing)) Then
        Me._valueCallbackAirQuality(Me, value)
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
    Public Overloads Function registerTimedReportCallback(callback As YAirQualityTimedReportCallback) As Integer
      Dim sensor As YSensor
      sensor = Me
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(sensor, True)
      Else
        YFunction._UpdateTimedReportCallbackList(sensor, False)
      End If
      Me._timedReportCallbackAirQuality = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackAirQuality Is Nothing)) Then
        Me._timedReportCallbackAirQuality(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of air quality sensors started using <c>yFirstAirQuality()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned air quality sensors order.
    '''   If you want to find a specific a air quality sensor, use <c>AirQuality.findAirQuality()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YAirQuality</c> object, corresponding to
    '''   a air quality sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are no more air quality sensors to enumerate.
    ''' </returns>
    '''/
    Public Function nextAirQuality() As YAirQuality
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YAirQuality.FindAirQuality(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of air quality sensors currently accessible.
    ''' <para>
    '''   Use the method <c>YAirQuality.nextAirQuality()</c> to iterate on
    '''   next air quality sensors.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YAirQuality</c> object, corresponding to
    '''   the first air quality sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstAirQuality() As YAirQuality
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("AirQuality", 0, p, size, neededsize, errmsg)
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
      Return YAirQuality.FindAirQuality(serial + "." + funcId)
    End Function

    REM --- (end of YAirQuality public methods declaration)

  End Class

  REM --- (YAirQuality functions)

  '''*
  ''' <summary>
  '''   Retrieves a air quality sensor for a given identifier.
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
  '''   This function does not require that the air quality sensor is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YAirQuality.isOnline()</c> to test if the air quality sensor is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a air quality sensor by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the air quality sensor, for instance
  '''   <c>MyDevice.airQuality</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YAirQuality</c> object allowing you to drive the air quality sensor.
  ''' </returns>
  '''/
  Public Function yFindAirQuality(ByVal func As String) As YAirQuality
    Return YAirQuality.FindAirQuality(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of air quality sensors currently accessible.
  ''' <para>
  '''   Use the method <c>YAirQuality.nextAirQuality()</c> to iterate on
  '''   next air quality sensors.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YAirQuality</c> object, corresponding to
  '''   the first air quality sensor currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstAirQuality() As YAirQuality
    Return YAirQuality.FirstAirQuality()
  End Function


  REM --- (end of YAirQuality functions)

End Module

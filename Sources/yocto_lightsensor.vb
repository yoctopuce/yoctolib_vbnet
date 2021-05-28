' ********************************************************************
'
'  $Id: yocto_lightsensor.vb 45292 2021-05-25 23:27:54Z mvuilleu $
'
'  Implements yFindLightSensor(), the high-level API for LightSensor functions
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

Module yocto_lightsensor

    REM --- (YLightSensor return codes)
    REM --- (end of YLightSensor return codes)
    REM --- (YLightSensor dlldef)
    REM --- (end of YLightSensor dlldef)
   REM --- (YLightSensor yapiwrapper)
   REM --- (end of YLightSensor yapiwrapper)
  REM --- (YLightSensor globals)

  Public Const Y_MEASURETYPE_HUMAN_EYE As Integer = 0
  Public Const Y_MEASURETYPE_WIDE_SPECTRUM As Integer = 1
  Public Const Y_MEASURETYPE_INFRARED As Integer = 2
  Public Const Y_MEASURETYPE_HIGH_RATE As Integer = 3
  Public Const Y_MEASURETYPE_HIGH_ENERGY As Integer = 4
  Public Const Y_MEASURETYPE_HIGH_RESOLUTION As Integer = 5
  Public Const Y_MEASURETYPE_INVALID As Integer = -1
  Public Delegate Sub YLightSensorValueCallback(ByVal func As YLightSensor, ByVal value As String)
  Public Delegate Sub YLightSensorTimedReportCallback(ByVal func As YLightSensor, ByVal measure As YMeasure)
  REM --- (end of YLightSensor globals)

  REM --- (YLightSensor class start)

  '''*
  ''' <summary>
  '''   The <c>YLightSensor</c> class allows you to read and configure Yoctopuce light sensors.
  ''' <para>
  '''   It inherits from <c>YSensor</c> class the core functions to read measurements,
  '''   to register callback functions, and to access the autonomous datalogger.
  '''   This class adds the ability to easily perform a one-point linear calibration
  '''   to compensate the effect of a glass or filter placed in front of the sensor.
  '''   For some light sensors with several working modes, this class can select the
  '''   desired working mode.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YLightSensor
    Inherits YSensor
    REM --- (end of YLightSensor class start)

    REM --- (YLightSensor definitions)
    Public Const MEASURETYPE_HUMAN_EYE As Integer = 0
    Public Const MEASURETYPE_WIDE_SPECTRUM As Integer = 1
    Public Const MEASURETYPE_INFRARED As Integer = 2
    Public Const MEASURETYPE_HIGH_RATE As Integer = 3
    Public Const MEASURETYPE_HIGH_ENERGY As Integer = 4
    Public Const MEASURETYPE_HIGH_RESOLUTION As Integer = 5
    Public Const MEASURETYPE_INVALID As Integer = -1
    REM --- (end of YLightSensor definitions)

    REM --- (YLightSensor attributes declaration)
    Protected _measureType As Integer
    Protected _valueCallbackLightSensor As YLightSensorValueCallback
    Protected _timedReportCallbackLightSensor As YLightSensorTimedReportCallback
    REM --- (end of YLightSensor attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "LightSensor"
      REM --- (YLightSensor attributes initialization)
      _measureType = MEASURETYPE_INVALID
      _valueCallbackLightSensor = Nothing
      _timedReportCallbackLightSensor = Nothing
      REM --- (end of YLightSensor attributes initialization)
    End Sub

    REM --- (YLightSensor private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("measureType") Then
        _measureType = CInt(json_val.getLong("measureType"))
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YLightSensor private methods declaration)

    REM --- (YLightSensor public methods declaration)

    Public Function set_currentValue(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("currentValue", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Changes the sensor-specific calibration parameter so that the current value
    '''   matches a desired target (linear scaling).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="calibratedVal">
    '''   the desired target value.
    ''' </param>
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function calibrate(ByVal calibratedVal As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(calibratedVal * 65536.0)))
      Return _setAttr("currentValue", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the type of light measure.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>YLightSensor.MEASURETYPE_HUMAN_EYE</c>, <c>YLightSensor.MEASURETYPE_WIDE_SPECTRUM</c>,
    '''   <c>YLightSensor.MEASURETYPE_INFRARED</c>, <c>YLightSensor.MEASURETYPE_HIGH_RATE</c>,
    '''   <c>YLightSensor.MEASURETYPE_HIGH_ENERGY</c> and <c>YLightSensor.MEASURETYPE_HIGH_RESOLUTION</c>
    '''   corresponding to the type of light measure
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YLightSensor.MEASURETYPE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_measureType() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return MEASURETYPE_INVALID
        End If
      End If
      res = Me._measureType
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the light sensor type used in the device.
    ''' <para>
    '''   The measure can either
    '''   approximate the response of the human eye, focus on a specific light
    '''   spectrum, depending on the capabilities of the light-sensitive cell.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>YLightSensor.MEASURETYPE_HUMAN_EYE</c>, <c>YLightSensor.MEASURETYPE_WIDE_SPECTRUM</c>,
    '''   <c>YLightSensor.MEASURETYPE_INFRARED</c>, <c>YLightSensor.MEASURETYPE_HIGH_RATE</c>,
    '''   <c>YLightSensor.MEASURETYPE_HIGH_ENERGY</c> and <c>YLightSensor.MEASURETYPE_HIGH_RESOLUTION</c>
    '''   corresponding to the light sensor type used in the device
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
    Public Function set_measureType(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("measureType", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves a light sensor for a given identifier.
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
    '''   This function does not require that the light sensor is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YLightSensor.isOnline()</c> to test if the light sensor is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a light sensor by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the light sensor, for instance
    '''   <c>LIGHTMK4.lightSensor</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YLightSensor</c> object allowing you to drive the light sensor.
    ''' </returns>
    '''/
    Public Shared Function FindLightSensor(func As String) As YLightSensor
      Dim obj As YLightSensor
      obj = CType(YFunction._FindFromCache("LightSensor", func), YLightSensor)
      If ((obj Is Nothing)) Then
        obj = New YLightSensor(func)
        YFunction._AddToCache("LightSensor", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YLightSensorValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackLightSensor = callback
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
      If (Not (Me._valueCallbackLightSensor Is Nothing)) Then
        Me._valueCallbackLightSensor(Me, value)
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
    Public Overloads Function registerTimedReportCallback(callback As YLightSensorTimedReportCallback) As Integer
      Dim sensor As YSensor
      sensor = Me
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(sensor, True)
      Else
        YFunction._UpdateTimedReportCallbackList(sensor, False)
      End If
      Me._timedReportCallbackLightSensor = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackLightSensor Is Nothing)) Then
        Me._timedReportCallbackLightSensor(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of light sensors started using <c>yFirstLightSensor()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned light sensors order.
    '''   If you want to find a specific a light sensor, use <c>LightSensor.findLightSensor()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YLightSensor</c> object, corresponding to
    '''   a light sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are no more light sensors to enumerate.
    ''' </returns>
    '''/
    Public Function nextLightSensor() As YLightSensor
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YLightSensor.FindLightSensor(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of light sensors currently accessible.
    ''' <para>
    '''   Use the method <c>YLightSensor.nextLightSensor()</c> to iterate on
    '''   next light sensors.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YLightSensor</c> object, corresponding to
    '''   the first light sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstLightSensor() As YLightSensor
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("LightSensor", 0, p, size, neededsize, errmsg)
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
      Return YLightSensor.FindLightSensor(serial + "." + funcId)
    End Function

    REM --- (end of YLightSensor public methods declaration)

  End Class

  REM --- (YLightSensor functions)

  '''*
  ''' <summary>
  '''   Retrieves a light sensor for a given identifier.
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
  '''   This function does not require that the light sensor is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YLightSensor.isOnline()</c> to test if the light sensor is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a light sensor by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the light sensor, for instance
  '''   <c>LIGHTMK4.lightSensor</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YLightSensor</c> object allowing you to drive the light sensor.
  ''' </returns>
  '''/
  Public Function yFindLightSensor(ByVal func As String) As YLightSensor
    Return YLightSensor.FindLightSensor(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of light sensors currently accessible.
  ''' <para>
  '''   Use the method <c>YLightSensor.nextLightSensor()</c> to iterate on
  '''   next light sensors.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YLightSensor</c> object, corresponding to
  '''   the first light sensor currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstLightSensor() As YLightSensor
    Return YLightSensor.FirstLightSensor()
  End Function


  REM --- (end of YLightSensor functions)

End Module

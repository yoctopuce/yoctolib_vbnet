'*********************************************************************
'*
'* $Id: yocto_compass.vb 28740 2017-10-03 08:09:13Z seb $
'*
'* Implements yFindCompass(), the high-level API for Compass functions
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

Module yocto_compass

    REM --- (YCompass return codes)
    REM --- (end of YCompass return codes)
    REM --- (YCompass dlldef)
    REM --- (end of YCompass dlldef)
  REM --- (YCompass globals)

  Public Const Y_BANDWIDTH_INVALID As Integer = YAPI.INVALID_INT
  Public Const Y_AXIS_X As Integer = 0
  Public Const Y_AXIS_Y As Integer = 1
  Public Const Y_AXIS_Z As Integer = 2
  Public Const Y_AXIS_INVALID As Integer = -1
  Public Const Y_MAGNETICHEADING_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Delegate Sub YCompassValueCallback(ByVal func As YCompass, ByVal value As String)
  Public Delegate Sub YCompassTimedReportCallback(ByVal func As YCompass, ByVal measure As YMeasure)
  REM --- (end of YCompass globals)

  REM --- (YCompass class start)

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
  Public Class YCompass
    Inherits YSensor
    REM --- (end of YCompass class start)

    REM --- (YCompass definitions)
    Public Const BANDWIDTH_INVALID As Integer = YAPI.INVALID_INT
    Public Const AXIS_X As Integer = 0
    Public Const AXIS_Y As Integer = 1
    Public Const AXIS_Z As Integer = 2
    Public Const AXIS_INVALID As Integer = -1
    Public Const MAGNETICHEADING_INVALID As Double = YAPI.INVALID_DOUBLE
    REM --- (end of YCompass definitions)

    REM --- (YCompass attributes declaration)
    Protected _bandwidth As Integer
    Protected _axis As Integer
    Protected _magneticHeading As Double
    Protected _valueCallbackCompass As YCompassValueCallback
    Protected _timedReportCallbackCompass As YCompassTimedReportCallback
    REM --- (end of YCompass attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Compass"
      REM --- (YCompass attributes initialization)
      _bandwidth = BANDWIDTH_INVALID
      _axis = AXIS_INVALID
      _magneticHeading = MAGNETICHEADING_INVALID
      _valueCallbackCompass = Nothing
      _timedReportCallbackCompass = Nothing
      REM --- (end of YCompass attributes initialization)
    End Sub

    REM --- (YCompass private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("bandwidth") Then
        _bandwidth = CInt(json_val.getLong("bandwidth"))
      End If
      If json_val.has("axis") Then
        _axis = CInt(json_val.getLong("axis"))
      End If
      If json_val.has("magneticHeading") Then
        _magneticHeading = Math.Round(json_val.getDouble("magneticHeading") * 1000.0 / 65536.0) / 1000.0
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YCompass private methods declaration)

    REM --- (YCompass public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the measure update frequency, measured in Hz (Yocto-3D-V2 only).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the measure update frequency, measured in Hz (Yocto-3D-V2 only)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_BANDWIDTH_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_bandwidth() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return BANDWIDTH_INVALID
        End If
      End If
      res = Me._bandwidth
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the measure update frequency, measured in Hz (Yocto-3D-V2 only).
    ''' <para>
    '''   When the
    '''   frequency is lower, the device performs averaging.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the measure update frequency, measured in Hz (Yocto-3D-V2 only)
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
    Public Function set_bandwidth(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("bandwidth", rest_val)
    End Function
    Public Function get_axis() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return AXIS_INVALID
        End If
      End If
      res = Me._axis
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the magnetic heading, regardless of the configured bearing.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the magnetic heading, regardless of the configured bearing
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_MAGNETICHEADING_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_magneticHeading() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return MAGNETICHEADING_INVALID
        End If
      End If
      res = Me._magneticHeading
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Retrieves a compass for a given identifier.
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
    '''   This function does not require that the compass is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YCompass.isOnline()</c> to test if the compass is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a compass by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the compass
    ''' </param>
    ''' <returns>
    '''   a <c>YCompass</c> object allowing you to drive the compass.
    ''' </returns>
    '''/
    Public Shared Function FindCompass(func As String) As YCompass
      Dim obj As YCompass
      obj = CType(YFunction._FindFromCache("Compass", func), YCompass)
      If ((obj Is Nothing)) Then
        obj = New YCompass(func)
        YFunction._AddToCache("Compass", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YCompassValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackCompass = callback
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
      If (Not (Me._valueCallbackCompass Is Nothing)) Then
        Me._valueCallbackCompass(Me, value)
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
    Public Overloads Function registerTimedReportCallback(callback As YCompassTimedReportCallback) As Integer
      Dim sensor As YSensor
      sensor = Me
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(sensor, True)
      Else
        YFunction._UpdateTimedReportCallbackList(sensor, False)
      End If
      Me._timedReportCallbackCompass = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackCompass Is Nothing)) Then
        Me._timedReportCallbackCompass(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of compasses started using <c>yFirstCompass()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YCompass</c> object, corresponding to
    '''   a compass currently online, or a <c>Nothing</c> pointer
    '''   if there are no more compasses to enumerate.
    ''' </returns>
    '''/
    Public Function nextCompass() As YCompass
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YCompass.FindCompass(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of compasses currently accessible.
    ''' <para>
    '''   Use the method <c>YCompass.nextCompass()</c> to iterate on
    '''   next compasses.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YCompass</c> object, corresponding to
    '''   the first compass currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstCompass() As YCompass
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Compass", 0, p, size, neededsize, errmsg)
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
      Return YCompass.FindCompass(serial + "." + funcId)
    End Function

    REM --- (end of YCompass public methods declaration)

  End Class

  REM --- (YCompass functions)

  '''*
  ''' <summary>
  '''   Retrieves a compass for a given identifier.
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
  '''   This function does not require that the compass is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YCompass.isOnline()</c> to test if the compass is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a compass by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the compass
  ''' </param>
  ''' <returns>
  '''   a <c>YCompass</c> object allowing you to drive the compass.
  ''' </returns>
  '''/
  Public Function yFindCompass(ByVal func As String) As YCompass
    Return YCompass.FindCompass(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of compasses currently accessible.
  ''' <para>
  '''   Use the method <c>YCompass.nextCompass()</c> to iterate on
  '''   next compasses.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YCompass</c> object, corresponding to
  '''   the first compass currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstCompass() As YCompass
    Return YCompass.FirstCompass()
  End Function


  REM --- (end of YCompass functions)

End Module

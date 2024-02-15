' ********************************************************************
'
'  $Id: svn_id $
'
'  Implements yFindThreshold(), the high-level API for Threshold functions
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

Module yocto_threshold

    REM --- (YThreshold return codes)
    REM --- (end of YThreshold return codes)
    REM --- (YThreshold dlldef)
    REM --- (end of YThreshold dlldef)
   REM --- (YThreshold yapiwrapper)
   REM --- (end of YThreshold yapiwrapper)
  REM --- (YThreshold globals)

  Public Const Y_THRESHOLDSTATE_SAFE As Integer = 0
  Public Const Y_THRESHOLDSTATE_ALERT As Integer = 1
  Public Const Y_THRESHOLDSTATE_INVALID As Integer = -1
  Public Const Y_TARGETSENSOR_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_ALERTLEVEL_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_SAFELEVEL_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Delegate Sub YThresholdValueCallback(ByVal func As YThreshold, ByVal value As String)
  Public Delegate Sub YThresholdTimedReportCallback(ByVal func As YThreshold, ByVal measure As YMeasure)
  REM --- (end of YThreshold globals)

  REM --- (YThreshold class start)

  '''*
  ''' <summary>
  '''   The <c>Threshold</c> class allows you define a threshold on a Yoctopuce sensor
  '''   to trigger a predefined action, on specific devices where this is implemented.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YThreshold
    Inherits YFunction
    REM --- (end of YThreshold class start)

    REM --- (YThreshold definitions)
    Public Const THRESHOLDSTATE_SAFE As Integer = 0
    Public Const THRESHOLDSTATE_ALERT As Integer = 1
    Public Const THRESHOLDSTATE_INVALID As Integer = -1
    Public Const TARGETSENSOR_INVALID As String = YAPI.INVALID_STRING
    Public Const ALERTLEVEL_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const SAFELEVEL_INVALID As Double = YAPI.INVALID_DOUBLE
    REM --- (end of YThreshold definitions)

    REM --- (YThreshold attributes declaration)
    Protected _thresholdState As Integer
    Protected _targetSensor As String
    Protected _alertLevel As Double
    Protected _safeLevel As Double
    Protected _valueCallbackThreshold As YThresholdValueCallback
    REM --- (end of YThreshold attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Threshold"
      REM --- (YThreshold attributes initialization)
      _thresholdState = THRESHOLDSTATE_INVALID
      _targetSensor = TARGETSENSOR_INVALID
      _alertLevel = ALERTLEVEL_INVALID
      _safeLevel = SAFELEVEL_INVALID
      _valueCallbackThreshold = Nothing
      REM --- (end of YThreshold attributes initialization)
    End Sub

    REM --- (YThreshold private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("thresholdState") Then
        _thresholdState = CInt(json_val.getLong("thresholdState"))
      End If
      If json_val.has("targetSensor") Then
        _targetSensor = json_val.getString("targetSensor")
      End If
      If json_val.has("alertLevel") Then
        _alertLevel = Math.Round(json_val.getDouble("alertLevel") / 65.536) / 1000.0
      End If
      If json_val.has("safeLevel") Then
        _safeLevel = Math.Round(json_val.getDouble("safeLevel") / 65.536) / 1000.0
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YThreshold private methods declaration)

    REM --- (YThreshold public methods declaration)
    '''*
    ''' <summary>
    '''   Returns current state of the threshold function.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>YThreshold.THRESHOLDSTATE_SAFE</c> or <c>YThreshold.THRESHOLDSTATE_ALERT</c>, according
    '''   to current state of the threshold function
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YThreshold.THRESHOLDSTATE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_thresholdState() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return THRESHOLDSTATE_INVALID
        End If
      End If
      res = Me._thresholdState
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the name of the sensor monitored by the threshold function.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the name of the sensor monitored by the threshold function
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YThreshold.TARGETSENSOR_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_targetSensor() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return TARGETSENSOR_INVALID
        End If
      End If
      res = Me._targetSensor
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the sensor alert level triggering the threshold function.
    ''' <para>
    '''   Remember to call the matching module <c>saveToFlash()</c>
    '''   method if you want to preserve the setting after reboot.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the sensor alert level triggering the threshold function
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
    Public Function set_alertLevel(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("alertLevel", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the sensor alert level, triggering the threshold function.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the sensor alert level, triggering the threshold function
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YThreshold.ALERTLEVEL_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_alertLevel() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return ALERTLEVEL_INVALID
        End If
      End If
      res = Me._alertLevel
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the sensor acceptable level for disabling the threshold function.
    ''' <para>
    '''   Remember to call the matching module <c>saveToFlash()</c>
    '''   method if you want to preserve the setting after reboot.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the sensor acceptable level for disabling the threshold function
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
    Public Function set_safeLevel(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("safeLevel", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the sensor acceptable level for disabling the threshold function.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the sensor acceptable level for disabling the threshold function
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YThreshold.SAFELEVEL_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_safeLevel() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return SAFELEVEL_INVALID
        End If
      End If
      res = Me._safeLevel
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Retrieves a threshold function for a given identifier.
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
    '''   This function does not require that the threshold function is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YThreshold.isOnline()</c> to test if the threshold function is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a threshold function by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the threshold function, for instance
    '''   <c>MyDevice.threshold1</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YThreshold</c> object allowing you to drive the threshold function.
    ''' </returns>
    '''/
    Public Shared Function FindThreshold(func As String) As YThreshold
      Dim obj As YThreshold
      obj = CType(YFunction._FindFromCache("Threshold", func), YThreshold)
      If ((obj Is Nothing)) Then
        obj = New YThreshold(func)
        YFunction._AddToCache("Threshold", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YThresholdValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackThreshold = callback
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
      If (Not (Me._valueCallbackThreshold Is Nothing)) Then
        Me._valueCallbackThreshold(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of threshold functions started using <c>yFirstThreshold()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned threshold functions order.
    '''   If you want to find a specific a threshold function, use <c>Threshold.findThreshold()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YThreshold</c> object, corresponding to
    '''   a threshold function currently online, or a <c>Nothing</c> pointer
    '''   if there are no more threshold functions to enumerate.
    ''' </returns>
    '''/
    Public Function nextThreshold() As YThreshold
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YThreshold.FindThreshold(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of threshold functions currently accessible.
    ''' <para>
    '''   Use the method <c>YThreshold.nextThreshold()</c> to iterate on
    '''   next threshold functions.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YThreshold</c> object, corresponding to
    '''   the first threshold function currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstThreshold() As YThreshold
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Threshold", 0, p, size, neededsize, errmsg)
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
      Return YThreshold.FindThreshold(serial + "." + funcId)
    End Function

    REM --- (end of YThreshold public methods declaration)

  End Class

  REM --- (YThreshold functions)

  '''*
  ''' <summary>
  '''   Retrieves a threshold function for a given identifier.
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
  '''   This function does not require that the threshold function is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YThreshold.isOnline()</c> to test if the threshold function is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a threshold function by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the threshold function, for instance
  '''   <c>MyDevice.threshold1</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YThreshold</c> object allowing you to drive the threshold function.
  ''' </returns>
  '''/
  Public Function yFindThreshold(ByVal func As String) As YThreshold
    Return YThreshold.FindThreshold(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of threshold functions currently accessible.
  ''' <para>
  '''   Use the method <c>YThreshold.nextThreshold()</c> to iterate on
  '''   next threshold functions.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YThreshold</c> object, corresponding to
  '''   the first threshold function currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstThreshold() As YThreshold
    Return YThreshold.FirstThreshold()
  End Function


  REM --- (end of YThreshold functions)

End Module

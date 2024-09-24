' ********************************************************************
'
'  $Id: yocto_multisenscontroller.vb 61964 2024-07-29 15:54:55Z seb $
'
'  Implements yFindMultiSensController(), the high-level API for MultiSensController functions
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

Module yocto_multisenscontroller

    REM --- (YMultiSensController return codes)
    REM --- (end of YMultiSensController return codes)
    REM --- (YMultiSensController dlldef)
    REM --- (end of YMultiSensController dlldef)
   REM --- (YMultiSensController yapiwrapper)
   REM --- (end of YMultiSensController yapiwrapper)
  REM --- (YMultiSensController globals)

  Public Const Y_NSENSORS_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_MAXSENSORS_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_MAINTENANCEMODE_FALSE As Integer = 0
  Public Const Y_MAINTENANCEMODE_TRUE As Integer = 1
  Public Const Y_MAINTENANCEMODE_INVALID As Integer = -1
  Public Const Y_LASTADDRESSDETECTED_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_COMMAND_INVALID As String = YAPI.INVALID_STRING
  Public Delegate Sub YMultiSensControllerValueCallback(ByVal func As YMultiSensController, ByVal value As String)
  Public Delegate Sub YMultiSensControllerTimedReportCallback(ByVal func As YMultiSensController, ByVal measure As YMeasure)
  REM --- (end of YMultiSensController globals)

  REM --- (YMultiSensController class start)

  '''*
  ''' <summary>
  '''   The <c>YMultiSensController</c> class allows you to set up a customized
  '''   sensor chain on devices featuring that functionality.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YMultiSensController
    Inherits YFunction
    REM --- (end of YMultiSensController class start)

    REM --- (YMultiSensController definitions)
    Public Const NSENSORS_INVALID As Integer = YAPI.INVALID_UINT
    Public Const MAXSENSORS_INVALID As Integer = YAPI.INVALID_UINT
    Public Const MAINTENANCEMODE_FALSE As Integer = 0
    Public Const MAINTENANCEMODE_TRUE As Integer = 1
    Public Const MAINTENANCEMODE_INVALID As Integer = -1
    Public Const LASTADDRESSDETECTED_INVALID As Integer = YAPI.INVALID_UINT
    Public Const COMMAND_INVALID As String = YAPI.INVALID_STRING
    REM --- (end of YMultiSensController definitions)

    REM --- (YMultiSensController attributes declaration)
    Protected _nSensors As Integer
    Protected _maxSensors As Integer
    Protected _maintenanceMode As Integer
    Protected _lastAddressDetected As Integer
    Protected _command As String
    Protected _valueCallbackMultiSensController As YMultiSensControllerValueCallback
    REM --- (end of YMultiSensController attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "MultiSensController"
      REM --- (YMultiSensController attributes initialization)
      _nSensors = NSENSORS_INVALID
      _maxSensors = MAXSENSORS_INVALID
      _maintenanceMode = MAINTENANCEMODE_INVALID
      _lastAddressDetected = LASTADDRESSDETECTED_INVALID
      _command = COMMAND_INVALID
      _valueCallbackMultiSensController = Nothing
      REM --- (end of YMultiSensController attributes initialization)
    End Sub

    REM --- (YMultiSensController private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("nSensors") Then
        _nSensors = CInt(json_val.getLong("nSensors"))
      End If
      If json_val.has("maxSensors") Then
        _maxSensors = CInt(json_val.getLong("maxSensors"))
      End If
      If json_val.has("maintenanceMode") Then
        If (json_val.getInt("maintenanceMode") > 0) Then _maintenanceMode = 1 Else _maintenanceMode = 0
      End If
      If json_val.has("lastAddressDetected") Then
        _lastAddressDetected = CInt(json_val.getLong("lastAddressDetected"))
      End If
      If json_val.has("command") Then
        _command = json_val.getString("command")
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YMultiSensController private methods declaration)

    REM --- (YMultiSensController public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the number of sensors to poll.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of sensors to poll
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMultiSensController.NSENSORS_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_nSensors() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return NSENSORS_INVALID
        End If
      End If
      res = Me._nSensors
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the number of sensors to poll.
    ''' <para>
    '''   Remember to call the
    '''   <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept. It is recommended to restart the
    '''   device with  <c>module->reboot()</c> after modifying
    '''   (and saving) this settings.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the number of sensors to poll
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
    Public Function set_nSensors(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("nSensors", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the maximum configurable sensor count allowed on this device.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the maximum configurable sensor count allowed on this device
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMultiSensController.MAXSENSORS_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_maxSensors() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return MAXSENSORS_INVALID
        End If
      End If
      res = Me._maxSensors
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns true when the device is in maintenance mode.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>YMultiSensController.MAINTENANCEMODE_FALSE</c> or <c>YMultiSensController.MAINTENANCEMODE_TRUE</c>,
    '''   according to true when the device is in maintenance mode
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMultiSensController.MAINTENANCEMODE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_maintenanceMode() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return MAINTENANCEMODE_INVALID
        End If
      End If
      res = Me._maintenanceMode
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the device mode to enable maintenance and to stop sensor polling.
    ''' <para>
    '''   This way, the device does not automatically restart when it cannot
    '''   communicate with one of the sensors.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>YMultiSensController.MAINTENANCEMODE_FALSE</c> or <c>YMultiSensController.MAINTENANCEMODE_TRUE</c>,
    '''   according to the device mode to enable maintenance and to stop sensor polling
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
    Public Function set_maintenanceMode(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("maintenanceMode", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the I2C address of the most recently detected sensor.
    ''' <para>
    '''   This method can
    '''   be used to in case of I2C communication error to determine what is the
    '''   last sensor that can be reached, or after a call to <c>setupAddress</c>
    '''   to make sure that the address change was properly processed.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the I2C address of the most recently detected sensor
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMultiSensController.LASTADDRESSDETECTED_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_lastAddressDetected() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return LASTADDRESSDETECTED_INVALID
        End If
      End If
      res = Me._lastAddressDetected
      Return res
    End Function

    Public Function get_command() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return COMMAND_INVALID
        End If
      End If
      res = Me._command
      Return res
    End Function


    Public Function set_command(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("command", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves a multi-sensor controller for a given identifier.
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
    '''   This function does not require that the multi-sensor controller is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YMultiSensController.isOnline()</c> to test if the multi-sensor controller is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a multi-sensor controller by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the multi-sensor controller, for instance
    '''   <c>YTEMPIR1.multiSensController</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YMultiSensController</c> object allowing you to drive the multi-sensor controller.
    ''' </returns>
    '''/
    Public Shared Function FindMultiSensController(func As String) As YMultiSensController
      Dim obj As YMultiSensController
      obj = CType(YFunction._FindFromCache("MultiSensController", func), YMultiSensController)
      If ((obj Is Nothing)) Then
        obj = New YMultiSensController(func)
        YFunction._AddToCache("MultiSensController", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YMultiSensControllerValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackMultiSensController = callback
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
      If (Not (Me._valueCallbackMultiSensController Is Nothing)) Then
        Me._valueCallbackMultiSensController(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Configures the I2C address of the only sensor connected to the device.
    ''' <para>
    '''   It is recommended to put the the device in maintenance mode before
    '''   changing sensor addresses.  This method is only intended to work with a single
    '''   sensor connected to the device. If several sensors are connected, the result
    '''   is unpredictable.
    ''' </para>
    ''' <para>
    '''   Note that the device is expecting to find a sensor or a string of sensors with specific
    '''   addresses. Check the device documentation to find out which addresses should be used.
    ''' </para>
    ''' </summary>
    ''' <param name="addr">
    '''   new address of the connected sensor
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function setupAddress(addr As Integer) As Integer
      Dim cmd As String
      Dim res As Integer = 0
      cmd = "A" + Convert.ToString(addr)
      res = Me.set_command(cmd)
      If Not(res = YAPI.SUCCESS) Then
        me._throw( YAPI.IO_ERROR,  "unable to trigger address change")
        return YAPI.IO_ERROR
      end if
      YAPI.Sleep(1500, Nothing)
      res = Me.get_lastAddressDetected()
      If Not(res > 0) Then
        me._throw( YAPI.IO_ERROR,  "IR sensor not found")
        return YAPI.IO_ERROR
      end if
      If Not(res = addr) Then
        me._throw( YAPI.IO_ERROR,  "address change failed")
        return YAPI.IO_ERROR
      end if
      Return YAPI.SUCCESS
    End Function

    '''*
    ''' <summary>
    '''   Triggers the I2C address detection procedure for the only sensor connected to the device.
    ''' <para>
    '''   This method is only intended to work with a single sensor connected to the device.
    '''   If several sensors are connected, the result is unpredictable.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the I2C address of the detected sensor, or 0 if none is found
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function get_sensorAddress() As Integer
      Dim res As Integer = 0
      res = Me.set_command("a")
      If Not(res = YAPI.SUCCESS) Then
        me._throw( YAPI.IO_ERROR,  "unable to trigger address detection")
        return res
      end if
      YAPI.Sleep(1000, Nothing)
      res = Me.get_lastAddressDetected()
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of multi-sensor controllers started using <c>yFirstMultiSensController()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned multi-sensor controllers order.
    '''   If you want to find a specific a multi-sensor controller, use
    '''   <c>MultiSensController.findMultiSensController()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YMultiSensController</c> object, corresponding to
    '''   a multi-sensor controller currently online, or a <c>Nothing</c> pointer
    '''   if there are no more multi-sensor controllers to enumerate.
    ''' </returns>
    '''/
    Public Function nextMultiSensController() As YMultiSensController
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YMultiSensController.FindMultiSensController(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of multi-sensor controllers currently accessible.
    ''' <para>
    '''   Use the method <c>YMultiSensController.nextMultiSensController()</c> to iterate on
    '''   next multi-sensor controllers.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YMultiSensController</c> object, corresponding to
    '''   the first multi-sensor controller currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstMultiSensController() As YMultiSensController
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("MultiSensController", 0, p, size, neededsize, errmsg)
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
      Return YMultiSensController.FindMultiSensController(serial + "." + funcId)
    End Function

    REM --- (end of YMultiSensController public methods declaration)

  End Class

  REM --- (YMultiSensController functions)

  '''*
  ''' <summary>
  '''   Retrieves a multi-sensor controller for a given identifier.
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
  '''   This function does not require that the multi-sensor controller is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YMultiSensController.isOnline()</c> to test if the multi-sensor controller is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a multi-sensor controller by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the multi-sensor controller, for instance
  '''   <c>YTEMPIR1.multiSensController</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YMultiSensController</c> object allowing you to drive the multi-sensor controller.
  ''' </returns>
  '''/
  Public Function yFindMultiSensController(ByVal func As String) As YMultiSensController
    Return YMultiSensController.FindMultiSensController(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of multi-sensor controllers currently accessible.
  ''' <para>
  '''   Use the method <c>YMultiSensController.nextMultiSensController()</c> to iterate on
  '''   next multi-sensor controllers.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YMultiSensController</c> object, corresponding to
  '''   the first multi-sensor controller currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstMultiSensController() As YMultiSensController
    Return YMultiSensController.FirstMultiSensController()
  End Function


  REM --- (end of YMultiSensController functions)

End Module

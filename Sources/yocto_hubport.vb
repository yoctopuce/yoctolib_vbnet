'*********************************************************************
'*
'* $Id: yocto_hubport.vb 31448 2018-08-08 09:13:11Z seb $
'*
'* Implements yFindHubPort(), the high-level API for HubPort functions
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

Module yocto_hubport

    REM --- (YHubPort return codes)
    REM --- (end of YHubPort return codes)
    REM --- (YHubPort dlldef)
    REM --- (end of YHubPort dlldef)
   REM --- (YHubPort yapiwrapper)
   REM --- (end of YHubPort yapiwrapper)
  REM --- (YHubPort globals)

  Public Const Y_ENABLED_FALSE As Integer = 0
  Public Const Y_ENABLED_TRUE As Integer = 1
  Public Const Y_ENABLED_INVALID As Integer = -1
  Public Const Y_PORTSTATE_OFF As Integer = 0
  Public Const Y_PORTSTATE_OVRLD As Integer = 1
  Public Const Y_PORTSTATE_ON As Integer = 2
  Public Const Y_PORTSTATE_RUN As Integer = 3
  Public Const Y_PORTSTATE_PROG As Integer = 4
  Public Const Y_PORTSTATE_INVALID As Integer = -1
  Public Const Y_BAUDRATE_INVALID As Integer = YAPI.INVALID_UINT
  Public Delegate Sub YHubPortValueCallback(ByVal func As YHubPort, ByVal value As String)
  Public Delegate Sub YHubPortTimedReportCallback(ByVal func As YHubPort, ByVal measure As YMeasure)
  REM --- (end of YHubPort globals)

  REM --- (YHubPort class start)

  '''*
  ''' <summary>
  '''   YHubPort objects provide control over the power supply for every
  '''   YoctoHub port and provide information about the device connected to it.
  ''' <para>
  '''   The logical name of a YHubPort is always automatically set to the
  '''   unique serial number of the Yoctopuce device connected to it.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YHubPort
    Inherits YFunction
    REM --- (end of YHubPort class start)

    REM --- (YHubPort definitions)
    Public Const ENABLED_FALSE As Integer = 0
    Public Const ENABLED_TRUE As Integer = 1
    Public Const ENABLED_INVALID As Integer = -1
    Public Const PORTSTATE_OFF As Integer = 0
    Public Const PORTSTATE_OVRLD As Integer = 1
    Public Const PORTSTATE_ON As Integer = 2
    Public Const PORTSTATE_RUN As Integer = 3
    Public Const PORTSTATE_PROG As Integer = 4
    Public Const PORTSTATE_INVALID As Integer = -1
    Public Const BAUDRATE_INVALID As Integer = YAPI.INVALID_UINT
    REM --- (end of YHubPort definitions)

    REM --- (YHubPort attributes declaration)
    Protected _enabled As Integer
    Protected _portState As Integer
    Protected _baudRate As Integer
    Protected _valueCallbackHubPort As YHubPortValueCallback
    REM --- (end of YHubPort attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "HubPort"
      REM --- (YHubPort attributes initialization)
      _enabled = ENABLED_INVALID
      _portState = PORTSTATE_INVALID
      _baudRate = BAUDRATE_INVALID
      _valueCallbackHubPort = Nothing
      REM --- (end of YHubPort attributes initialization)
    End Sub

    REM --- (YHubPort private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("enabled") Then
        If (json_val.getInt("enabled") > 0) Then _enabled = 1 Else _enabled = 0
      End If
      If json_val.has("portState") Then
        _portState = CInt(json_val.getLong("portState"))
      End If
      If json_val.has("baudRate") Then
        _baudRate = CInt(json_val.getLong("baudRate"))
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YHubPort private methods declaration)

    REM --- (YHubPort public methods declaration)
    '''*
    ''' <summary>
    '''   Returns true if the Yocto-hub port is powered, false otherwise.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_ENABLED_FALSE</c> or <c>Y_ENABLED_TRUE</c>, according to true if the Yocto-hub port is
    '''   powered, false otherwise
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
    '''   Changes the activation of the Yocto-hub port.
    ''' <para>
    '''   If the port is enabled, the
    '''   connected module is powered. Otherwise, port power is shut down.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>Y_ENABLED_FALSE</c> or <c>Y_ENABLED_TRUE</c>, according to the activation of the Yocto-hub port
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
    '''   Returns the current state of the Yocto-hub port.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_PORTSTATE_OFF</c>, <c>Y_PORTSTATE_OVRLD</c>, <c>Y_PORTSTATE_ON</c>,
    '''   <c>Y_PORTSTATE_RUN</c> and <c>Y_PORTSTATE_PROG</c> corresponding to the current state of the Yocto-hub port
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_PORTSTATE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_portState() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return PORTSTATE_INVALID
        End If
      End If
      res = Me._portState
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the current baud rate used by this Yocto-hub port, in kbps.
    ''' <para>
    '''   The default value is 1000 kbps, but a slower rate may be used if communication
    '''   problems are encountered.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the current baud rate used by this Yocto-hub port, in kbps
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_BAUDRATE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_baudRate() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return BAUDRATE_INVALID
        End If
      End If
      res = Me._baudRate
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Retrieves a Yocto-hub port for a given identifier.
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
    '''   This function does not require that the Yocto-hub port is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YHubPort.isOnline()</c> to test if the Yocto-hub port is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a Yocto-hub port by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the Yocto-hub port
    ''' </param>
    ''' <returns>
    '''   a <c>YHubPort</c> object allowing you to drive the Yocto-hub port.
    ''' </returns>
    '''/
    Public Shared Function FindHubPort(func As String) As YHubPort
      Dim obj As YHubPort
      obj = CType(YFunction._FindFromCache("HubPort", func), YHubPort)
      If ((obj Is Nothing)) Then
        obj = New YHubPort(func)
        YFunction._AddToCache("HubPort", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YHubPortValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackHubPort = callback
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
      If (Not (Me._valueCallbackHubPort Is Nothing)) Then
        Me._valueCallbackHubPort(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of Yocto-hub ports started using <c>yFirstHubPort()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YHubPort</c> object, corresponding to
    '''   a Yocto-hub port currently online, or a <c>Nothing</c> pointer
    '''   if there are no more Yocto-hub ports to enumerate.
    ''' </returns>
    '''/
    Public Function nextHubPort() As YHubPort
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YHubPort.FindHubPort(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of Yocto-hub ports currently accessible.
    ''' <para>
    '''   Use the method <c>YHubPort.nextHubPort()</c> to iterate on
    '''   next Yocto-hub ports.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YHubPort</c> object, corresponding to
    '''   the first Yocto-hub port currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstHubPort() As YHubPort
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("HubPort", 0, p, size, neededsize, errmsg)
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
      Return YHubPort.FindHubPort(serial + "." + funcId)
    End Function

    REM --- (end of YHubPort public methods declaration)

  End Class

  REM --- (YHubPort functions)

  '''*
  ''' <summary>
  '''   Retrieves a Yocto-hub port for a given identifier.
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
  '''   This function does not require that the Yocto-hub port is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YHubPort.isOnline()</c> to test if the Yocto-hub port is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a Yocto-hub port by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the Yocto-hub port
  ''' </param>
  ''' <returns>
  '''   a <c>YHubPort</c> object allowing you to drive the Yocto-hub port.
  ''' </returns>
  '''/
  Public Function yFindHubPort(ByVal func As String) As YHubPort
    Return YHubPort.FindHubPort(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of Yocto-hub ports currently accessible.
  ''' <para>
  '''   Use the method <c>YHubPort.nextHubPort()</c> to iterate on
  '''   next Yocto-hub ports.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YHubPort</c> object, corresponding to
  '''   the first Yocto-hub port currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstHubPort() As YHubPort
    Return YHubPort.FirstHubPort()
  End Function


  REM --- (end of YHubPort functions)

End Module

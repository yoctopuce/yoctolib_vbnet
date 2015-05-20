'*********************************************************************
'*
'* $Id: yocto_bluetoothlink.vb 20325 2015-05-12 15:34:50Z seb $
'*
'* Implements yFindBluetoothLink(), the high-level API for BluetoothLink functions
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

Module yocto_bluetoothlink

    REM --- (YBluetoothLink return codes)
    REM --- (end of YBluetoothLink return codes)
    REM --- (YBluetoothLink dlldef)
    REM --- (end of YBluetoothLink dlldef)
  REM --- (YBluetoothLink globals)

  Public Const Y_OWNADDRESS_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_PAIRINGPIN_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_REMOTEADDRESS_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_MESSAGE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_COMMAND_INVALID As String = YAPI.INVALID_STRING
  Public Delegate Sub YBluetoothLinkValueCallback(ByVal func As YBluetoothLink, ByVal value As String)
  Public Delegate Sub YBluetoothLinkTimedReportCallback(ByVal func As YBluetoothLink, ByVal measure As YMeasure)
  REM --- (end of YBluetoothLink globals)

  REM --- (YBluetoothLink class start)

  '''*
  ''' <summary>
  '''   BluetoothLink function provides control over bluetooth link
  '''   and status for devices that are bluetooth-enabled.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YBluetoothLink
    Inherits YFunction
    REM --- (end of YBluetoothLink class start)

    REM --- (YBluetoothLink definitions)
    Public Const OWNADDRESS_INVALID As String = YAPI.INVALID_STRING
    Public Const PAIRINGPIN_INVALID As String = YAPI.INVALID_STRING
    Public Const REMOTEADDRESS_INVALID As String = YAPI.INVALID_STRING
    Public Const MESSAGE_INVALID As String = YAPI.INVALID_STRING
    Public Const COMMAND_INVALID As String = YAPI.INVALID_STRING
    REM --- (end of YBluetoothLink definitions)

    REM --- (YBluetoothLink attributes declaration)
    Protected _ownAddress As String
    Protected _pairingPin As String
    Protected _remoteAddress As String
    Protected _message As String
    Protected _command As String
    Protected _valueCallbackBluetoothLink As YBluetoothLinkValueCallback
    REM --- (end of YBluetoothLink attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "BluetoothLink"
      REM --- (YBluetoothLink attributes initialization)
      _ownAddress = OWNADDRESS_INVALID
      _pairingPin = PAIRINGPIN_INVALID
      _remoteAddress = REMOTEADDRESS_INVALID
      _message = MESSAGE_INVALID
      _command = COMMAND_INVALID
      _valueCallbackBluetoothLink = Nothing
      REM --- (end of YBluetoothLink attributes initialization)
    End Sub

    REM --- (YBluetoothLink private methods declaration)

    Protected Overrides Function _parseAttr(ByRef member As TJSONRECORD) As Integer
      If (member.name = "ownAddress") Then
        _ownAddress = member.svalue
        Return 1
      End If
      If (member.name = "pairingPin") Then
        _pairingPin = member.svalue
        Return 1
      End If
      If (member.name = "remoteAddress") Then
        _remoteAddress = member.svalue
        Return 1
      End If
      If (member.name = "message") Then
        _message = member.svalue
        Return 1
      End If
      If (member.name = "command") Then
        _command = member.svalue
        Return 1
      End If
      Return MyBase._parseAttr(member)
    End Function

    REM --- (end of YBluetoothLink private methods declaration)

    REM --- (YBluetoothLink public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the MAC-48 address of the bluetooth interface, which is unique on the bluetooth network.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the MAC-48 address of the bluetooth interface, which is unique on the
    '''   bluetooth network
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_OWNADDRESS_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_ownAddress() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return OWNADDRESS_INVALID
        End If
      End If
      Return Me._ownAddress
    End Function

    '''*
    ''' <summary>
    '''   Returns an opaque string if a PIN code has been configured in the device to access
    '''   the SIM card, or an empty string if none has been configured or if the code provided
    '''   was rejected by the SIM card.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to an opaque string if a PIN code has been configured in the device to access
    '''   the SIM card, or an empty string if none has been configured or if the code provided
    '''   was rejected by the SIM card
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_PAIRINGPIN_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_pairingPin() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return PAIRINGPIN_INVALID
        End If
      End If
      Return Me._pairingPin
    End Function


    '''*
    ''' <summary>
    '''   Changes the PIN code used by the module for bluetooth pairing.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method of the module to save the
    '''   new value in the device flash.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the PIN code used by the module for bluetooth pairing
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
    Public Function set_pairingPin(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("pairingPin", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the MAC-48 address of the remote device to connect to.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the MAC-48 address of the remote device to connect to
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_REMOTEADDRESS_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_remoteAddress() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return REMOTEADDRESS_INVALID
        End If
      End If
      Return Me._remoteAddress
    End Function


    '''*
    ''' <summary>
    '''   Changes the MAC-48 address defining which remote device to connect to.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the MAC-48 address defining which remote device to connect to
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
    Public Function set_remoteAddress(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("remoteAddress", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the latest status message from the bluetooth interface.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the latest status message from the bluetooth interface
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_MESSAGE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_message() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return MESSAGE_INVALID
        End If
      End If
      Return Me._message
    End Function

    Public Function get_command() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return COMMAND_INVALID
        End If
      End If
      Return Me._command
    End Function


    Public Function set_command(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("command", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves a cellular interface for a given identifier.
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
    '''   This function does not require that the cellular interface is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YBluetoothLink.isOnline()</c> to test if the cellular interface is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a cellular interface by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the cellular interface
    ''' </param>
    ''' <returns>
    '''   a <c>YBluetoothLink</c> object allowing you to drive the cellular interface.
    ''' </returns>
    '''/
    Public Shared Function FindBluetoothLink(func As String) As YBluetoothLink
      Dim obj As YBluetoothLink
      obj = CType(YFunction._FindFromCache("BluetoothLink", func), YBluetoothLink)
      If ((obj Is Nothing)) Then
        obj = New YBluetoothLink(func)
        YFunction._AddToCache("BluetoothLink", func, obj)
      End If
      Return obj
    End Function

    '''*
    ''' <summary>
    '''   Registers the callback function that is invoked on every change of advertised value.
    ''' <para>
    '''   The callback is invoked only during the execution of <c>ySleep</c> or <c>yHandleEvents</c>.
    '''   This provides control over the time when the callback is triggered. For good responsiveness, remember to call
    '''   one of these two functions periodically. To unregister a callback, pass a null pointer as argument.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="callback">
    '''   the callback function to call, or a null pointer. The callback function should take two
    '''   arguments: the function object of which the value has changed, and the character string describing
    '''   the new advertised value.
    ''' @noreturn
    ''' </param>
    '''/
    Public Overloads Function registerValueCallback(callback As YBluetoothLinkValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackBluetoothLink = callback
      REM // Immediately invoke value callback with current value
      If (Not (callback Is Nothing) And Me.isOnline()) Then
        val = Me._advertisedValue
        If (Not (val = "")) Then
          Me._invokeValueCallback(val)
        End If
      End If
      Return 0
    End Function

    Public Overrides Function _invokeValueCallback(value As String) As Integer
      If (Not (Me._valueCallbackBluetoothLink Is Nothing)) Then
        Me._valueCallbackBluetoothLink(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Attempt to connect to the previously selected remote device.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function connect() As Integer
      Return Me.set_command("C")
    End Function

    '''*
    ''' <summary>
    '''   Disconnect from the previously selected remote device.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function disconnect() As Integer
      Return Me.set_command("D")
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of cellular interfaces started using <c>yFirstBluetoothLink()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YBluetoothLink</c> object, corresponding to
    '''   a cellular interface currently online, or a <c>null</c> pointer
    '''   if there are no more cellular interfaces to enumerate.
    ''' </returns>
    '''/
    Public Function nextBluetoothLink() As YBluetoothLink
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YBluetoothLink.FindBluetoothLink(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of cellular interfaces currently accessible.
    ''' <para>
    '''   Use the method <c>YBluetoothLink.nextBluetoothLink()</c> to iterate on
    '''   next cellular interfaces.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YBluetoothLink</c> object, corresponding to
    '''   the first cellular interface currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstBluetoothLink() As YBluetoothLink
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("BluetoothLink", 0, p, size, neededsize, errmsg)
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
      Return YBluetoothLink.FindBluetoothLink(serial + "." + funcId)
    End Function

    REM --- (end of YBluetoothLink public methods declaration)

  End Class

  REM --- (BluetoothLink functions)

  '''*
  ''' <summary>
  '''   Retrieves a cellular interface for a given identifier.
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
  '''   This function does not require that the cellular interface is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YBluetoothLink.isOnline()</c> to test if the cellular interface is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a cellular interface by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the cellular interface
  ''' </param>
  ''' <returns>
  '''   a <c>YBluetoothLink</c> object allowing you to drive the cellular interface.
  ''' </returns>
  '''/
  Public Function yFindBluetoothLink(ByVal func As String) As YBluetoothLink
    Return YBluetoothLink.FindBluetoothLink(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of cellular interfaces currently accessible.
  ''' <para>
  '''   Use the method <c>YBluetoothLink.nextBluetoothLink()</c> to iterate on
  '''   next cellular interfaces.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YBluetoothLink</c> object, corresponding to
  '''   the first cellular interface currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstBluetoothLink() As YBluetoothLink
    Return YBluetoothLink.FirstBluetoothLink()
  End Function


  REM --- (end of BluetoothLink functions)

End Module

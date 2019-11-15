' ********************************************************************
'
'  $Id: yocto_bluetoothlink.vb 37827 2019-10-25 13:07:48Z mvuilleu $
'
'  Implements yFindBluetoothLink(), the high-level API for BluetoothLink functions
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

Module yocto_bluetoothlink

    REM --- (YBluetoothLink return codes)
    REM --- (end of YBluetoothLink return codes)
    REM --- (YBluetoothLink dlldef)
    REM --- (end of YBluetoothLink dlldef)
   REM --- (YBluetoothLink yapiwrapper)
   REM --- (end of YBluetoothLink yapiwrapper)
  REM --- (YBluetoothLink globals)

  Public Const Y_OWNADDRESS_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_PAIRINGPIN_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_REMOTEADDRESS_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_REMOTENAME_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_MUTE_FALSE As Integer = 0
  Public Const Y_MUTE_TRUE As Integer = 1
  Public Const Y_MUTE_INVALID As Integer = -1
  Public Const Y_PREAMPLIFIER_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_VOLUME_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_LINKSTATE_DOWN As Integer = 0
  Public Const Y_LINKSTATE_FREE As Integer = 1
  Public Const Y_LINKSTATE_SEARCH As Integer = 2
  Public Const Y_LINKSTATE_EXISTS As Integer = 3
  Public Const Y_LINKSTATE_LINKED As Integer = 4
  Public Const Y_LINKSTATE_PLAY As Integer = 5
  Public Const Y_LINKSTATE_INVALID As Integer = -1
  Public Const Y_LINKQUALITY_INVALID As Integer = YAPI.INVALID_UINT
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
    Public Const REMOTENAME_INVALID As String = YAPI.INVALID_STRING
    Public Const MUTE_FALSE As Integer = 0
    Public Const MUTE_TRUE As Integer = 1
    Public Const MUTE_INVALID As Integer = -1
    Public Const PREAMPLIFIER_INVALID As Integer = YAPI.INVALID_UINT
    Public Const VOLUME_INVALID As Integer = YAPI.INVALID_UINT
    Public Const LINKSTATE_DOWN As Integer = 0
    Public Const LINKSTATE_FREE As Integer = 1
    Public Const LINKSTATE_SEARCH As Integer = 2
    Public Const LINKSTATE_EXISTS As Integer = 3
    Public Const LINKSTATE_LINKED As Integer = 4
    Public Const LINKSTATE_PLAY As Integer = 5
    Public Const LINKSTATE_INVALID As Integer = -1
    Public Const LINKQUALITY_INVALID As Integer = YAPI.INVALID_UINT
    Public Const COMMAND_INVALID As String = YAPI.INVALID_STRING
    REM --- (end of YBluetoothLink definitions)

    REM --- (YBluetoothLink attributes declaration)
    Protected _ownAddress As String
    Protected _pairingPin As String
    Protected _remoteAddress As String
    Protected _remoteName As String
    Protected _mute As Integer
    Protected _preAmplifier As Integer
    Protected _volume As Integer
    Protected _linkState As Integer
    Protected _linkQuality As Integer
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
      _remoteName = REMOTENAME_INVALID
      _mute = MUTE_INVALID
      _preAmplifier = PREAMPLIFIER_INVALID
      _volume = VOLUME_INVALID
      _linkState = LINKSTATE_INVALID
      _linkQuality = LINKQUALITY_INVALID
      _command = COMMAND_INVALID
      _valueCallbackBluetoothLink = Nothing
      REM --- (end of YBluetoothLink attributes initialization)
    End Sub

    REM --- (YBluetoothLink private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("ownAddress") Then
        _ownAddress = json_val.getString("ownAddress")
      End If
      If json_val.has("pairingPin") Then
        _pairingPin = json_val.getString("pairingPin")
      End If
      If json_val.has("remoteAddress") Then
        _remoteAddress = json_val.getString("remoteAddress")
      End If
      If json_val.has("remoteName") Then
        _remoteName = json_val.getString("remoteName")
      End If
      If json_val.has("mute") Then
        If (json_val.getInt("mute") > 0) Then _mute = 1 Else _mute = 0
      End If
      If json_val.has("preAmplifier") Then
        _preAmplifier = CInt(json_val.getLong("preAmplifier"))
      End If
      If json_val.has("volume") Then
        _volume = CInt(json_val.getLong("volume"))
      End If
      If json_val.has("linkState") Then
        _linkState = CInt(json_val.getLong("linkState"))
      End If
      If json_val.has("linkQuality") Then
        _linkQuality = CInt(json_val.getLong("linkQuality"))
      End If
      If json_val.has("command") Then
        _command = json_val.getString("command")
      End If
      Return MyBase._parseAttr(json_val)
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
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return OWNADDRESS_INVALID
        End If
      End If
      res = Me._ownAddress
      Return res
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
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return PAIRINGPIN_INVALID
        End If
      End If
      res = Me._pairingPin
      Return res
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
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return REMOTEADDRESS_INVALID
        End If
      End If
      res = Me._remoteAddress
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the MAC-48 address defining which remote device to connect to.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c>
    '''   method of the module if the modification must be kept.
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
    '''   Returns the bluetooth name the remote device, if found on the bluetooth network.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the bluetooth name the remote device, if found on the bluetooth network
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_REMOTENAME_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_remoteName() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return REMOTENAME_INVALID
        End If
      End If
      res = Me._remoteName
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the state of the mute function.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_MUTE_FALSE</c> or <c>Y_MUTE_TRUE</c>, according to the state of the mute function
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_MUTE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_mute() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return MUTE_INVALID
        End If
      End If
      res = Me._mute
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the state of the mute function.
    ''' <para>
    '''   Remember to call the matching module
    '''   <c>saveToFlash()</c> method to save the setting permanently.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>Y_MUTE_FALSE</c> or <c>Y_MUTE_TRUE</c>, according to the state of the mute function
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
    Public Function set_mute(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("mute", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the audio pre-amplifier volume, in per cents.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the audio pre-amplifier volume, in per cents
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_PREAMPLIFIER_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_preAmplifier() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return PREAMPLIFIER_INVALID
        End If
      End If
      res = Me._preAmplifier
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the audio pre-amplifier volume, in per cents.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c>
    '''   method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the audio pre-amplifier volume, in per cents
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
    Public Function set_preAmplifier(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("preAmplifier", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the connected headset volume, in per cents.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the connected headset volume, in per cents
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_VOLUME_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_volume() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return VOLUME_INVALID
        End If
      End If
      res = Me._volume
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the connected headset volume, in per cents.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the connected headset volume, in per cents
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
    Public Function set_volume(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("volume", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the bluetooth link state.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_LINKSTATE_DOWN</c>, <c>Y_LINKSTATE_FREE</c>, <c>Y_LINKSTATE_SEARCH</c>,
    '''   <c>Y_LINKSTATE_EXISTS</c>, <c>Y_LINKSTATE_LINKED</c> and <c>Y_LINKSTATE_PLAY</c> corresponding to
    '''   the bluetooth link state
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_LINKSTATE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_linkState() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return LINKSTATE_INVALID
        End If
      End If
      res = Me._linkState
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the bluetooth receiver signal strength, in pourcents, or 0 if no connection is established.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the bluetooth receiver signal strength, in pourcents, or 0 if no
    '''   connection is established
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_LINKQUALITY_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_linkQuality() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return LINKQUALITY_INVALID
        End If
      End If
      res = Me._linkQuality
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
    ''' <para>
    '''   If a call to this object's is_online() method returns FALSE although
    '''   you are certain that the matching device is plugged, make sure that you did
    '''   call registerHub() at application initialization time.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the cellular interface, for instance
    '''   <c>MyDevice.bluetoothLink1</c>.
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
    Public Overloads Function registerValueCallback(callback As YBluetoothLinkValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackBluetoothLink = callback
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
    '''   Caution: You can't make any assumption about the returned cellular interfaces order.
    '''   If you want to find a specific a cellular interface, use <c>BluetoothLink.findBluetoothLink()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YBluetoothLink</c> object, corresponding to
    '''   a cellular interface currently online, or a <c>Nothing</c> pointer
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
    '''   the first cellular interface currently online, or a <c>Nothing</c> pointer
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

  REM --- (YBluetoothLink functions)

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
  ''' <para>
  '''   If a call to this object's is_online() method returns FALSE although
  '''   you are certain that the matching device is plugged, make sure that you did
  '''   call registerHub() at application initialization time.
  ''' </para>
  ''' <para>
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the cellular interface, for instance
  '''   <c>MyDevice.bluetoothLink1</c>.
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
  '''   the first cellular interface currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstBluetoothLink() As YBluetoothLink
    Return YBluetoothLink.FirstBluetoothLink()
  End Function


  REM --- (end of YBluetoothLink functions)

End Module

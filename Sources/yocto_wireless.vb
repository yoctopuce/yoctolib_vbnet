'*********************************************************************
'*
'* $Id: yocto_wireless.vb 27240 2017-04-24 12:26:37Z seb $
'*
'* Implements yFindWireless(), the high-level API for Wireless functions
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

Module yocto_wireless

  REM --- (generated code: YWlanRecord globals)

  REM --- (end of generated code: YWlanRecord globals)

  REM --- (generated code: YWireless globals)

  Public Const Y_LINKQUALITY_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_SSID_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_CHANNEL_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_SECURITY_UNKNOWN As Integer = 0
  Public Const Y_SECURITY_OPEN As Integer = 1
  Public Const Y_SECURITY_WEP As Integer = 2
  Public Const Y_SECURITY_WPA As Integer = 3
  Public Const Y_SECURITY_WPA2 As Integer = 4
  Public Const Y_SECURITY_INVALID As Integer = -1
  Public Const Y_MESSAGE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_WLANCONFIG_INVALID As String = YAPI.INVALID_STRING
  Public Delegate Sub YWirelessValueCallback(ByVal func As YWireless, ByVal value As String)
  Public Delegate Sub YWirelessTimedReportCallback(ByVal func As YWireless, ByVal measure As YMeasure)
  REM --- (end of generated code: YWireless globals)


  REM --- (generated code: YWlanRecord class start)

  Public Class YWlanRecord
    REM --- (end of generated code: YWlanRecord class start)
    REM --- (generated code: YWlanRecord definitions)
    REM --- (end of generated code: YWlanRecord definitions)
    REM --- (generated code: YWlanRecord attributes)
    Protected _ssid As String
    Protected _channel As Integer
    Protected _sec As String
    Protected _rssi As Integer
    REM --- (end of generated code: YWlanRecord attributes)

    REM --- (generated code: YWlanRecord private methods declaration)

    REM --- (end of generated code: YWlanRecord private methods declaration)

    REM --- (generated code: YWlanRecord public methods declaration)
    Public Overridable Function get_ssid() As String
      Return Me._ssid
    End Function

    Public Overridable Function get_channel() As Integer
      Return Me._channel
    End Function

    Public Overridable Function get_security() As String
      Return Me._sec
    End Function

    Public Overridable Function get_linkQuality() As Integer
      Return Me._rssi
    End Function



    REM --- (end of generated code: YWlanRecord public methods declaration)

  

    Public Sub New(ByVal data As String)
      Dim obj As YJSONObject  = New YJSONObject(data)
      obj.parse()
      Me._ssid = obj.getString("ssid")
      Me._sec =obj.getString("sec")
      Me._channel = CInt(obj.getInt("channel"))
      Me._rssi = CInt(obj.getInt("rssi"))     
    End Sub

  End Class
  




  
  REM --- (generated code: YWireless class start)

  '''*
  ''' <summary>
  '''   YWireless functions provides control over wireless network parameters
  '''   and status for devices that are wireless-enabled.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YWireless
    Inherits YFunction
    REM --- (end of generated code: YWireless class start)
    REM --- (generated code: YWireless definitions)
    Public Const LINKQUALITY_INVALID As Integer = YAPI.INVALID_UINT
    Public Const SSID_INVALID As String = YAPI.INVALID_STRING
    Public Const CHANNEL_INVALID As Integer = YAPI.INVALID_UINT
    Public Const SECURITY_UNKNOWN As Integer = 0
    Public Const SECURITY_OPEN As Integer = 1
    Public Const SECURITY_WEP As Integer = 2
    Public Const SECURITY_WPA As Integer = 3
    Public Const SECURITY_WPA2 As Integer = 4
    Public Const SECURITY_INVALID As Integer = -1
    Public Const MESSAGE_INVALID As String = YAPI.INVALID_STRING
    Public Const WLANCONFIG_INVALID As String = YAPI.INVALID_STRING
    REM --- (end of generated code: YWireless definitions)


    REM --- (generated code: YWireless attributes declaration)
    Protected _linkQuality As Integer
    Protected _ssid As String
    Protected _channel As Integer
    Protected _security As Integer
    Protected _message As String
    Protected _wlanConfig As String
    Protected _valueCallbackWireless As YWirelessValueCallback
    REM --- (end of generated code: YWireless attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.new(func)
      _className = "Wireless"
      REM --- (generated code: YWireless attributes initialization)
      _linkQuality = LINKQUALITY_INVALID
      _ssid = SSID_INVALID
      _channel = CHANNEL_INVALID
      _security = SECURITY_INVALID
      _message = MESSAGE_INVALID
      _wlanConfig = WLANCONFIG_INVALID
      _valueCallbackWireless = Nothing
      REM --- (end of generated code: YWireless attributes initialization)
    End Sub

      REM --- (generated code: YWireless private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("linkQuality") Then
        _linkQuality = CInt(json_val.getLong("linkQuality"))
      End If
      If json_val.has("ssid") Then
        _ssid = json_val.getString("ssid")
      End If
      If json_val.has("channel") Then
        _channel = CInt(json_val.getLong("channel"))
      End If
      If json_val.has("security") Then
        _security = CInt(json_val.getLong("security"))
      End If
      If json_val.has("message") Then
        _message = json_val.getString("message")
      End If
      If json_val.has("wlanConfig") Then
        _wlanConfig = json_val.getString("wlanConfig")
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of generated code: YWireless private methods declaration)

    REM --- (generated code: YWireless public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the link quality, expressed in percent.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the link quality, expressed in percent
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_LINKQUALITY_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_linkQuality() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return LINKQUALITY_INVALID
        End If
      End If
      res = Me._linkQuality
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the wireless network name (SSID).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the wireless network name (SSID)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_SSID_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_ssid() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return SSID_INVALID
        End If
      End If
      res = Me._ssid
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the 802.11 channel currently used, or 0 when the selected network has not been found.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the 802.11 channel currently used, or 0 when the selected network has not been found
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_CHANNEL_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_channel() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return CHANNEL_INVALID
        End If
      End If
      res = Me._channel
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the security algorithm used by the selected wireless network.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_SECURITY_UNKNOWN</c>, <c>Y_SECURITY_OPEN</c>, <c>Y_SECURITY_WEP</c>,
    '''   <c>Y_SECURITY_WPA</c> and <c>Y_SECURITY_WPA2</c> corresponding to the security algorithm used by
    '''   the selected wireless network
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_SECURITY_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_security() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return SECURITY_INVALID
        End If
      End If
      res = Me._security
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the latest status message from the wireless interface.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the latest status message from the wireless interface
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_MESSAGE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_message() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return MESSAGE_INVALID
        End If
      End If
      res = Me._message
      Return res
    End Function

    Public Function get_wlanConfig() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return WLANCONFIG_INVALID
        End If
      End If
      res = Me._wlanConfig
      Return res
    End Function


    Public Function set_wlanConfig(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("wlanConfig", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves a wireless lan interface for a given identifier.
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
    '''   This function does not require that the wireless lan interface is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YWireless.isOnline()</c> to test if the wireless lan interface is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a wireless lan interface by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the wireless lan interface
    ''' </param>
    ''' <returns>
    '''   a <c>YWireless</c> object allowing you to drive the wireless lan interface.
    ''' </returns>
    '''/
    Public Shared Function FindWireless(func As String) As YWireless
      Dim obj As YWireless
      obj = CType(YFunction._FindFromCache("Wireless", func), YWireless)
      If ((obj Is Nothing)) Then
        obj = New YWireless(func)
        YFunction._AddToCache("Wireless", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YWirelessValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackWireless = callback
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
      If (Not (Me._valueCallbackWireless Is Nothing)) Then
        Me._valueCallbackWireless(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Changes the configuration of the wireless lan interface to connect to an existing
    '''   access point (infrastructure mode).
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method and then to reboot the module to apply this setting.
    ''' </para>
    ''' </summary>
    ''' <param name="ssid">
    '''   the name of the network to connect to
    ''' </param>
    ''' <param name="securityKey">
    '''   the network key, as a character string
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function joinNetwork(ssid As String, securityKey As String) As Integer
      Return Me.set_wlanConfig("INFRA:" +  ssid + "\" + securityKey)
    End Function

    '''*
    ''' <summary>
    '''   Changes the configuration of the wireless lan interface to create an ad-hoc
    '''   wireless network, without using an access point.
    ''' <para>
    '''   On the YoctoHub-Wireless-g,
    '''   it is best to use softAPNetworkInstead(), which emulates an access point
    '''   (Soft AP) which is more efficient and more widely supported than ad-hoc networks.
    ''' </para>
    ''' <para>
    '''   When a security key is specified for an ad-hoc network, the network is protected
    '''   by a WEP40 key (5 characters or 10 hexadecimal digits) or WEP128 key (13 characters
    '''   or 26 hexadecimal digits). It is recommended to use a well-randomized WEP128 key
    '''   using 26 hexadecimal digits to maximize security.
    '''   Remember to call the <c>saveToFlash()</c> method and then to reboot the module
    '''   to apply this setting.
    ''' </para>
    ''' </summary>
    ''' <param name="ssid">
    '''   the name of the network to connect to
    ''' </param>
    ''' <param name="securityKey">
    '''   the network key, as a character string
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function adhocNetwork(ssid As String, securityKey As String) As Integer
      Return Me.set_wlanConfig("ADHOC:" +  ssid + "\" + securityKey)
    End Function

    '''*
    ''' <summary>
    '''   Changes the configuration of the wireless lan interface to create a new wireless
    '''   network by emulating a WiFi access point (Soft AP).
    ''' <para>
    '''   This function can only be
    '''   used with the YoctoHub-Wireless-g.
    ''' </para>
    ''' <para>
    '''   When a security key is specified for a SoftAP network, the network is protected
    '''   by a WEP40 key (5 characters or 10 hexadecimal digits) or WEP128 key (13 characters
    '''   or 26 hexadecimal digits). It is recommended to use a well-randomized WEP128 key
    '''   using 26 hexadecimal digits to maximize security.
    '''   Remember to call the <c>saveToFlash()</c> method and then to reboot the module to apply this setting.
    ''' </para>
    ''' </summary>
    ''' <param name="ssid">
    '''   the name of the network to connect to
    ''' </param>
    ''' <param name="securityKey">
    '''   the network key, as a character string
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function softAPNetwork(ssid As String, securityKey As String) As Integer
      Return Me.set_wlanConfig("SOFTAP:" +  ssid + "\" + securityKey)
    End Function

    '''*
    ''' <summary>
    '''   Returns a list of YWlanRecord objects that describe detected Wireless networks.
    ''' <para>
    '''   This list is not updated when the module is already connected to an acces point (infrastructure mode).
    '''   To force an update of this list, <c>adhocNetwork()</c> must be called to disconnect
    '''   the module from the current network. The returned list must be unallocated by the caller.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a list of <c>YWlanRecord</c> objects, containing the SSID, channel,
    '''   link quality and the type of security of the wireless network.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty list.
    ''' </para>
    '''/
    Public Overridable Function get_detectedWlans() As List(Of YWlanRecord)
      Dim i_i As Integer
      Dim json As Byte()
      Dim wlanlist As List(Of String) = New List(Of String)()
      Dim res As List(Of YWlanRecord) = New List(Of YWlanRecord)()
      
      json = Me._download("wlan.json?by=name")
      wlanlist = Me._json_get_array(json)
      res.Clear()
      For i_i = 0 To wlanlist.Count - 1
        res.Add(New YWlanRecord(wlanlist(i_i)))
      Next i_i
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of wireless lan interfaces started using <c>yFirstWireless()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YWireless</c> object, corresponding to
    '''   a wireless lan interface currently online, or a <c>Nothing</c> pointer
    '''   if there are no more wireless lan interfaces to enumerate.
    ''' </returns>
    '''/
    Public Function nextWireless() As YWireless
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YWireless.FindWireless(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of wireless lan interfaces currently accessible.
    ''' <para>
    '''   Use the method <c>YWireless.nextWireless()</c> to iterate on
    '''   next wireless lan interfaces.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YWireless</c> object, corresponding to
    '''   the first wireless lan interface currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstWireless() As YWireless
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Wireless", 0, p, size, neededsize, errmsg)
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
      Return YWireless.FindWireless(serial + "." + funcId)
    End Function

    REM --- (end of generated code: YWireless public methods declaration)

  End Class

  REM --- (generated code: Wireless functions)

  '''*
  ''' <summary>
  '''   Retrieves a wireless lan interface for a given identifier.
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
  '''   This function does not require that the wireless lan interface is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YWireless.isOnline()</c> to test if the wireless lan interface is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a wireless lan interface by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the wireless lan interface
  ''' </param>
  ''' <returns>
  '''   a <c>YWireless</c> object allowing you to drive the wireless lan interface.
  ''' </returns>
  '''/
  Public Function yFindWireless(ByVal func As String) As YWireless
    Return YWireless.FindWireless(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of wireless lan interfaces currently accessible.
  ''' <para>
  '''   Use the method <c>YWireless.nextWireless()</c> to iterate on
  '''   next wireless lan interfaces.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YWireless</c> object, corresponding to
  '''   the first wireless lan interface currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstWireless() As YWireless
    Return YWireless.FirstWireless()
  End Function


  REM --- (end of generated code: Wireless functions)

End Module

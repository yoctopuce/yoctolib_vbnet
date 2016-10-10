'*********************************************************************
'*
'* $Id: yocto_network.vb 25275 2016-08-24 13:42:24Z mvuilleu $
'*
'* Implements yFindNetwork(), the high-level API for Network functions
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

Module yocto_network

    REM --- (YNetwork return codes)
    REM --- (end of YNetwork return codes)
    REM --- (YNetwork dlldef)
    REM --- (end of YNetwork dlldef)
  REM --- (YNetwork globals)

  Public Const Y_READINESS_DOWN As Integer = 0
  Public Const Y_READINESS_EXISTS As Integer = 1
  Public Const Y_READINESS_LINKED As Integer = 2
  Public Const Y_READINESS_LAN_OK As Integer = 3
  Public Const Y_READINESS_WWW_OK As Integer = 4
  Public Const Y_READINESS_INVALID As Integer = -1
  Public Const Y_MACADDRESS_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_IPADDRESS_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_SUBNETMASK_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_ROUTER_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_IPCONFIG_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_PRIMARYDNS_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_SECONDARYDNS_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_NTPSERVER_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_USERPASSWORD_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_ADMINPASSWORD_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_HTTPPORT_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_DEFAULTPAGE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_DISCOVERABLE_FALSE As Integer = 0
  Public Const Y_DISCOVERABLE_TRUE As Integer = 1
  Public Const Y_DISCOVERABLE_INVALID As Integer = -1
  Public Const Y_WWWWATCHDOGDELAY_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_CALLBACKURL_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_CALLBACKMETHOD_POST As Integer = 0
  Public Const Y_CALLBACKMETHOD_GET As Integer = 1
  Public Const Y_CALLBACKMETHOD_PUT As Integer = 2
  Public Const Y_CALLBACKMETHOD_INVALID As Integer = -1
  Public Const Y_CALLBACKENCODING_FORM As Integer = 0
  Public Const Y_CALLBACKENCODING_JSON As Integer = 1
  Public Const Y_CALLBACKENCODING_JSON_ARRAY As Integer = 2
  Public Const Y_CALLBACKENCODING_CSV As Integer = 3
  Public Const Y_CALLBACKENCODING_YOCTO_API As Integer = 4
  Public Const Y_CALLBACKENCODING_JSON_NUM As Integer = 5
  Public Const Y_CALLBACKENCODING_EMONCMS As Integer = 6
  Public Const Y_CALLBACKENCODING_AZURE As Integer = 7
  Public Const Y_CALLBACKENCODING_INFLUXDB As Integer = 8
  Public Const Y_CALLBACKENCODING_MQTT As Integer = 9
  Public Const Y_CALLBACKENCODING_INVALID As Integer = -1
  Public Const Y_CALLBACKCREDENTIALS_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_CALLBACKINITIALDELAY_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_CALLBACKMINDELAY_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_CALLBACKMAXDELAY_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_POECURRENT_INVALID As Integer = YAPI.INVALID_UINT
  Public Delegate Sub YNetworkValueCallback(ByVal func As YNetwork, ByVal value As String)
  Public Delegate Sub YNetworkTimedReportCallback(ByVal func As YNetwork, ByVal measure As YMeasure)
  REM --- (end of YNetwork globals)

  REM --- (YNetwork class start)

  '''*
  ''' <summary>
  '''   YNetwork objects provide access to TCP/IP parameters of Yoctopuce
  '''   modules that include a built-in network interface.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YNetwork
    Inherits YFunction
    REM --- (end of YNetwork class start)

    REM --- (YNetwork definitions)
    Public Const READINESS_DOWN As Integer = 0
    Public Const READINESS_EXISTS As Integer = 1
    Public Const READINESS_LINKED As Integer = 2
    Public Const READINESS_LAN_OK As Integer = 3
    Public Const READINESS_WWW_OK As Integer = 4
    Public Const READINESS_INVALID As Integer = -1
    Public Const MACADDRESS_INVALID As String = YAPI.INVALID_STRING
    Public Const IPADDRESS_INVALID As String = YAPI.INVALID_STRING
    Public Const SUBNETMASK_INVALID As String = YAPI.INVALID_STRING
    Public Const ROUTER_INVALID As String = YAPI.INVALID_STRING
    Public Const IPCONFIG_INVALID As String = YAPI.INVALID_STRING
    Public Const PRIMARYDNS_INVALID As String = YAPI.INVALID_STRING
    Public Const SECONDARYDNS_INVALID As String = YAPI.INVALID_STRING
    Public Const NTPSERVER_INVALID As String = YAPI.INVALID_STRING
    Public Const USERPASSWORD_INVALID As String = YAPI.INVALID_STRING
    Public Const ADMINPASSWORD_INVALID As String = YAPI.INVALID_STRING
    Public Const HTTPPORT_INVALID As Integer = YAPI.INVALID_UINT
    Public Const DEFAULTPAGE_INVALID As String = YAPI.INVALID_STRING
    Public Const DISCOVERABLE_FALSE As Integer = 0
    Public Const DISCOVERABLE_TRUE As Integer = 1
    Public Const DISCOVERABLE_INVALID As Integer = -1
    Public Const WWWWATCHDOGDELAY_INVALID As Integer = YAPI.INVALID_UINT
    Public Const CALLBACKURL_INVALID As String = YAPI.INVALID_STRING
    Public Const CALLBACKMETHOD_POST As Integer = 0
    Public Const CALLBACKMETHOD_GET As Integer = 1
    Public Const CALLBACKMETHOD_PUT As Integer = 2
    Public Const CALLBACKMETHOD_INVALID As Integer = -1
    Public Const CALLBACKENCODING_FORM As Integer = 0
    Public Const CALLBACKENCODING_JSON As Integer = 1
    Public Const CALLBACKENCODING_JSON_ARRAY As Integer = 2
    Public Const CALLBACKENCODING_CSV As Integer = 3
    Public Const CALLBACKENCODING_YOCTO_API As Integer = 4
    Public Const CALLBACKENCODING_JSON_NUM As Integer = 5
    Public Const CALLBACKENCODING_EMONCMS As Integer = 6
    Public Const CALLBACKENCODING_AZURE As Integer = 7
    Public Const CALLBACKENCODING_INFLUXDB As Integer = 8
    Public Const CALLBACKENCODING_MQTT As Integer = 9
    Public Const CALLBACKENCODING_INVALID As Integer = -1
    Public Const CALLBACKCREDENTIALS_INVALID As String = YAPI.INVALID_STRING
    Public Const CALLBACKINITIALDELAY_INVALID As Integer = YAPI.INVALID_UINT
    Public Const CALLBACKMINDELAY_INVALID As Integer = YAPI.INVALID_UINT
    Public Const CALLBACKMAXDELAY_INVALID As Integer = YAPI.INVALID_UINT
    Public Const POECURRENT_INVALID As Integer = YAPI.INVALID_UINT
    REM --- (end of YNetwork definitions)

    REM --- (YNetwork attributes declaration)
    Protected _readiness As Integer
    Protected _macAddress As String
    Protected _ipAddress As String
    Protected _subnetMask As String
    Protected _router As String
    Protected _ipConfig As String
    Protected _primaryDNS As String
    Protected _secondaryDNS As String
    Protected _ntpServer As String
    Protected _userPassword As String
    Protected _adminPassword As String
    Protected _httpPort As Integer
    Protected _defaultPage As String
    Protected _discoverable As Integer
    Protected _wwwWatchdogDelay As Integer
    Protected _callbackUrl As String
    Protected _callbackMethod As Integer
    Protected _callbackEncoding As Integer
    Protected _callbackCredentials As String
    Protected _callbackInitialDelay As Integer
    Protected _callbackMinDelay As Integer
    Protected _callbackMaxDelay As Integer
    Protected _poeCurrent As Integer
    Protected _valueCallbackNetwork As YNetworkValueCallback
    REM --- (end of YNetwork attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Network"
      REM --- (YNetwork attributes initialization)
      _readiness = READINESS_INVALID
      _macAddress = MACADDRESS_INVALID
      _ipAddress = IPADDRESS_INVALID
      _subnetMask = SUBNETMASK_INVALID
      _router = ROUTER_INVALID
      _ipConfig = IPCONFIG_INVALID
      _primaryDNS = PRIMARYDNS_INVALID
      _secondaryDNS = SECONDARYDNS_INVALID
      _ntpServer = NTPSERVER_INVALID
      _userPassword = USERPASSWORD_INVALID
      _adminPassword = ADMINPASSWORD_INVALID
      _httpPort = HTTPPORT_INVALID
      _defaultPage = DEFAULTPAGE_INVALID
      _discoverable = DISCOVERABLE_INVALID
      _wwwWatchdogDelay = WWWWATCHDOGDELAY_INVALID
      _callbackUrl = CALLBACKURL_INVALID
      _callbackMethod = CALLBACKMETHOD_INVALID
      _callbackEncoding = CALLBACKENCODING_INVALID
      _callbackCredentials = CALLBACKCREDENTIALS_INVALID
      _callbackInitialDelay = CALLBACKINITIALDELAY_INVALID
      _callbackMinDelay = CALLBACKMINDELAY_INVALID
      _callbackMaxDelay = CALLBACKMAXDELAY_INVALID
      _poeCurrent = POECURRENT_INVALID
      _valueCallbackNetwork = Nothing
      REM --- (end of YNetwork attributes initialization)
    End Sub

    REM --- (YNetwork private methods declaration)

    Protected Overrides Function _parseAttr(ByRef member As TJSONRECORD) As Integer
      If (member.name = "readiness") Then
        _readiness = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "macAddress") Then
        _macAddress = member.svalue
        Return 1
      End If
      If (member.name = "ipAddress") Then
        _ipAddress = member.svalue
        Return 1
      End If
      If (member.name = "subnetMask") Then
        _subnetMask = member.svalue
        Return 1
      End If
      If (member.name = "router") Then
        _router = member.svalue
        Return 1
      End If
      If (member.name = "ipConfig") Then
        _ipConfig = member.svalue
        Return 1
      End If
      If (member.name = "primaryDNS") Then
        _primaryDNS = member.svalue
        Return 1
      End If
      If (member.name = "secondaryDNS") Then
        _secondaryDNS = member.svalue
        Return 1
      End If
      If (member.name = "ntpServer") Then
        _ntpServer = member.svalue
        Return 1
      End If
      If (member.name = "userPassword") Then
        _userPassword = member.svalue
        Return 1
      End If
      If (member.name = "adminPassword") Then
        _adminPassword = member.svalue
        Return 1
      End If
      If (member.name = "httpPort") Then
        _httpPort = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "defaultPage") Then
        _defaultPage = member.svalue
        Return 1
      End If
      If (member.name = "discoverable") Then
        If (member.ivalue > 0) Then _discoverable = 1 Else _discoverable = 0
        Return 1
      End If
      If (member.name = "wwwWatchdogDelay") Then
        _wwwWatchdogDelay = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "callbackUrl") Then
        _callbackUrl = member.svalue
        Return 1
      End If
      If (member.name = "callbackMethod") Then
        _callbackMethod = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "callbackEncoding") Then
        _callbackEncoding = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "callbackCredentials") Then
        _callbackCredentials = member.svalue
        Return 1
      End If
      If (member.name = "callbackInitialDelay") Then
        _callbackInitialDelay = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "callbackMinDelay") Then
        _callbackMinDelay = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "callbackMaxDelay") Then
        _callbackMaxDelay = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "poeCurrent") Then
        _poeCurrent = CInt(member.ivalue)
        Return 1
      End If
      Return MyBase._parseAttr(member)
    End Function

    REM --- (end of YNetwork private methods declaration)

    REM --- (YNetwork public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the current established working mode of the network interface.
    ''' <para>
    '''   Level zero (DOWN_0) means that no hardware link has been detected. Either there is no signal
    '''   on the network cable, or the selected wireless access point cannot be detected.
    '''   Level 1 (LIVE_1) is reached when the network is detected, but is not yet connected.
    '''   For a wireless network, this shows that the requested SSID is present.
    '''   Level 2 (LINK_2) is reached when the hardware connection is established.
    '''   For a wired network connection, level 2 means that the cable is attached at both ends.
    '''   For a connection to a wireless access point, it shows that the security parameters
    '''   are properly configured. For an ad-hoc wireless connection, it means that there is
    '''   at least one other device connected on the ad-hoc network.
    '''   Level 3 (DHCP_3) is reached when an IP address has been obtained using DHCP.
    '''   Level 4 (DNS_4) is reached when the DNS server is reachable on the network.
    '''   Level 5 (WWW_5) is reached when global connectivity is demonstrated by properly loading the
    '''   current time from an NTP server.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_READINESS_DOWN</c>, <c>Y_READINESS_EXISTS</c>, <c>Y_READINESS_LINKED</c>,
    '''   <c>Y_READINESS_LAN_OK</c> and <c>Y_READINESS_WWW_OK</c> corresponding to the current established
    '''   working mode of the network interface
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_READINESS_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_readiness() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return READINESS_INVALID
        End If
      End If
      Return Me._readiness
    End Function

    '''*
    ''' <summary>
    '''   Returns the MAC address of the network interface.
    ''' <para>
    '''   The MAC address is also available on a sticker
    '''   on the module, in both numeric and barcode forms.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the MAC address of the network interface
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_MACADDRESS_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_macAddress() As String
      If (Me._cacheExpiration = 0) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return MACADDRESS_INVALID
        End If
      End If
      Return Me._macAddress
    End Function

    '''*
    ''' <summary>
    '''   Returns the IP address currently in use by the device.
    ''' <para>
    '''   The address may have been configured
    '''   statically, or provided by a DHCP server.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the IP address currently in use by the device
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_IPADDRESS_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_ipAddress() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return IPADDRESS_INVALID
        End If
      End If
      Return Me._ipAddress
    End Function

    '''*
    ''' <summary>
    '''   Returns the subnet mask currently used by the device.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the subnet mask currently used by the device
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_SUBNETMASK_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_subnetMask() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return SUBNETMASK_INVALID
        End If
      End If
      Return Me._subnetMask
    End Function

    '''*
    ''' <summary>
    '''   Returns the IP address of the router on the device subnet (default gateway).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the IP address of the router on the device subnet (default gateway)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_ROUTER_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_router() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return ROUTER_INVALID
        End If
      End If
      Return Me._router
    End Function

    Public Function get_ipConfig() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return IPCONFIG_INVALID
        End If
      End If
      Return Me._ipConfig
    End Function


    Public Function set_ipConfig(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("ipConfig", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the IP address of the primary name server to be used by the module.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the IP address of the primary name server to be used by the module
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_PRIMARYDNS_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_primaryDNS() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return PRIMARYDNS_INVALID
        End If
      End If
      Return Me._primaryDNS
    End Function


    '''*
    ''' <summary>
    '''   Changes the IP address of the primary name server to be used by the module.
    ''' <para>
    '''   When using DHCP, if a value is specified, it overrides the value received from the DHCP server.
    '''   Remember to call the <c>saveToFlash()</c> method and then to reboot the module to apply this setting.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the IP address of the primary name server to be used by the module
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
    Public Function set_primaryDNS(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("primaryDNS", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the IP address of the secondary name server to be used by the module.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the IP address of the secondary name server to be used by the module
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_SECONDARYDNS_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_secondaryDNS() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return SECONDARYDNS_INVALID
        End If
      End If
      Return Me._secondaryDNS
    End Function


    '''*
    ''' <summary>
    '''   Changes the IP address of the secondary name server to be used by the module.
    ''' <para>
    '''   When using DHCP, if a value is specified, it overrides the value received from the DHCP server.
    '''   Remember to call the <c>saveToFlash()</c> method and then to reboot the module to apply this setting.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the IP address of the secondary name server to be used by the module
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
    Public Function set_secondaryDNS(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("secondaryDNS", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the IP address of the NTP server to be used by the device.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the IP address of the NTP server to be used by the device
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_NTPSERVER_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_ntpServer() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return NTPSERVER_INVALID
        End If
      End If
      Return Me._ntpServer
    End Function


    '''*
    ''' <summary>
    '''   Changes the IP address of the NTP server to be used by the module.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method and then to reboot the module to apply this setting.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the IP address of the NTP server to be used by the module
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
    Public Function set_ntpServer(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("ntpServer", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns a hash string if a password has been set for "user" user,
    '''   or an empty string otherwise.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to a hash string if a password has been set for "user" user,
    '''   or an empty string otherwise
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_USERPASSWORD_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_userPassword() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return USERPASSWORD_INVALID
        End If
      End If
      Return Me._userPassword
    End Function


    '''*
    ''' <summary>
    '''   Changes the password for the "user" user.
    ''' <para>
    '''   This password becomes instantly required
    '''   to perform any use of the module. If the specified value is an
    '''   empty string, a password is not required anymore.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the password for the "user" user
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
    Public Function set_userPassword(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("userPassword", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns a hash string if a password has been set for user "admin",
    '''   or an empty string otherwise.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to a hash string if a password has been set for user "admin",
    '''   or an empty string otherwise
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_ADMINPASSWORD_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_adminPassword() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return ADMINPASSWORD_INVALID
        End If
      End If
      Return Me._adminPassword
    End Function


    '''*
    ''' <summary>
    '''   Changes the password for the "admin" user.
    ''' <para>
    '''   This password becomes instantly required
    '''   to perform any change of the module state. If the specified value is an
    '''   empty string, a password is not required anymore.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the password for the "admin" user
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
    Public Function set_adminPassword(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("adminPassword", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the HTML page to serve for the URL "/"" of the hub.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the HTML page to serve for the URL "/"" of the hub
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_HTTPPORT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_httpPort() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return HTTPPORT_INVALID
        End If
      End If
      Return Me._httpPort
    End Function


    '''*
    ''' <summary>
    '''   Changes the default HTML page returned by the hub.
    ''' <para>
    '''   If not value are set the hub return
    '''   "index.html" which is the web interface of the hub. It is possible de change this page
    '''   for file that has been uploaded on the hub.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the default HTML page returned by the hub
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
    Public Function set_httpPort(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("httpPort", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the HTML page to serve for the URL "/"" of the hub.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the HTML page to serve for the URL "/"" of the hub
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_DEFAULTPAGE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_defaultPage() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return DEFAULTPAGE_INVALID
        End If
      End If
      Return Me._defaultPage
    End Function


    '''*
    ''' <summary>
    '''   Changes the default HTML page returned by the hub.
    ''' <para>
    '''   If not value are set the hub return
    '''   "index.html" which is the web interface of the hub. It is possible de change this page
    '''   for file that has been uploaded on the hub.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the default HTML page returned by the hub
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
    Public Function set_defaultPage(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("defaultPage", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the activation state of the multicast announce protocols to allow easy
    '''   discovery of the module in the network neighborhood (uPnP/Bonjour protocol).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_DISCOVERABLE_FALSE</c> or <c>Y_DISCOVERABLE_TRUE</c>, according to the activation state
    '''   of the multicast announce protocols to allow easy
    '''   discovery of the module in the network neighborhood (uPnP/Bonjour protocol)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_DISCOVERABLE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_discoverable() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return DISCOVERABLE_INVALID
        End If
      End If
      Return Me._discoverable
    End Function


    '''*
    ''' <summary>
    '''   Changes the activation state of the multicast announce protocols to allow easy
    '''   discovery of the module in the network neighborhood (uPnP/Bonjour protocol).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>Y_DISCOVERABLE_FALSE</c> or <c>Y_DISCOVERABLE_TRUE</c>, according to the activation state
    '''   of the multicast announce protocols to allow easy
    '''   discovery of the module in the network neighborhood (uPnP/Bonjour protocol)
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
    Public Function set_discoverable(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("discoverable", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the allowed downtime of the WWW link (in seconds) before triggering an automated
    '''   reboot to try to recover Internet connectivity.
    ''' <para>
    '''   A zero value disables automated reboot
    '''   in case of Internet connectivity loss.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the allowed downtime of the WWW link (in seconds) before triggering an automated
    '''   reboot to try to recover Internet connectivity
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_WWWWATCHDOGDELAY_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_wwwWatchdogDelay() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return WWWWATCHDOGDELAY_INVALID
        End If
      End If
      Return Me._wwwWatchdogDelay
    End Function


    '''*
    ''' <summary>
    '''   Changes the allowed downtime of the WWW link (in seconds) before triggering an automated
    '''   reboot to try to recover Internet connectivity.
    ''' <para>
    '''   A zero value disables automated reboot
    '''   in case of Internet connectivity loss. The smallest valid non-zero timeout is
    '''   90 seconds.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the allowed downtime of the WWW link (in seconds) before triggering an automated
    '''   reboot to try to recover Internet connectivity
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
    Public Function set_wwwWatchdogDelay(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("wwwWatchdogDelay", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the callback URL to notify of significant state changes.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the callback URL to notify of significant state changes
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_CALLBACKURL_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_callbackUrl() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return CALLBACKURL_INVALID
        End If
      End If
      Return Me._callbackUrl
    End Function


    '''*
    ''' <summary>
    '''   Changes the callback URL to notify significant state changes.
    ''' <para>
    '''   Remember to call the
    '''   <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the callback URL to notify significant state changes
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
    Public Function set_callbackUrl(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("callbackUrl", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the HTTP method used to notify callbacks for significant state changes.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_CALLBACKMETHOD_POST</c>, <c>Y_CALLBACKMETHOD_GET</c> and
    '''   <c>Y_CALLBACKMETHOD_PUT</c> corresponding to the HTTP method used to notify callbacks for
    '''   significant state changes
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_CALLBACKMETHOD_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_callbackMethod() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return CALLBACKMETHOD_INVALID
        End If
      End If
      Return Me._callbackMethod
    End Function


    '''*
    ''' <summary>
    '''   Changes the HTTP method used to notify callbacks for significant state changes.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>Y_CALLBACKMETHOD_POST</c>, <c>Y_CALLBACKMETHOD_GET</c> and
    '''   <c>Y_CALLBACKMETHOD_PUT</c> corresponding to the HTTP method used to notify callbacks for
    '''   significant state changes
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
    Public Function set_callbackMethod(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("callbackMethod", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the encoding standard to use for representing notification values.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_CALLBACKENCODING_FORM</c>, <c>Y_CALLBACKENCODING_JSON</c>,
    '''   <c>Y_CALLBACKENCODING_JSON_ARRAY</c>, <c>Y_CALLBACKENCODING_CSV</c>,
    '''   <c>Y_CALLBACKENCODING_YOCTO_API</c>, <c>Y_CALLBACKENCODING_JSON_NUM</c>,
    '''   <c>Y_CALLBACKENCODING_EMONCMS</c>, <c>Y_CALLBACKENCODING_AZURE</c>,
    '''   <c>Y_CALLBACKENCODING_INFLUXDB</c> and <c>Y_CALLBACKENCODING_MQTT</c> corresponding to the encoding
    '''   standard to use for representing notification values
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_CALLBACKENCODING_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_callbackEncoding() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return CALLBACKENCODING_INVALID
        End If
      End If
      Return Me._callbackEncoding
    End Function


    '''*
    ''' <summary>
    '''   Changes the encoding standard to use for representing notification values.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>Y_CALLBACKENCODING_FORM</c>, <c>Y_CALLBACKENCODING_JSON</c>,
    '''   <c>Y_CALLBACKENCODING_JSON_ARRAY</c>, <c>Y_CALLBACKENCODING_CSV</c>,
    '''   <c>Y_CALLBACKENCODING_YOCTO_API</c>, <c>Y_CALLBACKENCODING_JSON_NUM</c>,
    '''   <c>Y_CALLBACKENCODING_EMONCMS</c>, <c>Y_CALLBACKENCODING_AZURE</c>,
    '''   <c>Y_CALLBACKENCODING_INFLUXDB</c> and <c>Y_CALLBACKENCODING_MQTT</c> corresponding to the encoding
    '''   standard to use for representing notification values
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
    Public Function set_callbackEncoding(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("callbackEncoding", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns a hashed version of the notification callback credentials if set,
    '''   or an empty string otherwise.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to a hashed version of the notification callback credentials if set,
    '''   or an empty string otherwise
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_CALLBACKCREDENTIALS_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_callbackCredentials() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return CALLBACKCREDENTIALS_INVALID
        End If
      End If
      Return Me._callbackCredentials
    End Function


    '''*
    ''' <summary>
    '''   Changes the credentials required to connect to the callback address.
    ''' <para>
    '''   The credentials
    '''   must be provided as returned by function <c>get_callbackCredentials</c>,
    '''   in the form <c>username:hash</c>. The method used to compute the hash varies according
    '''   to the the authentication scheme implemented by the callback, For Basic authentication,
    '''   the hash is the MD5 of the string <c>username:password</c>. For Digest authentication,
    '''   the hash is the MD5 of the string <c>username:realm:password</c>. For a simpler
    '''   way to configure callback credentials, use function <c>callbackLogin</c> instead.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the credentials required to connect to the callback address
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
    Public Function set_callbackCredentials(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("callbackCredentials", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Connects to the notification callback and saves the credentials required to
    '''   log into it.
    ''' <para>
    '''   The password is not stored into the module, only a hashed
    '''   copy of the credentials are saved. Remember to call the
    '''   <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="username">
    '''   username required to log to the callback
    ''' </param>
    ''' <param name="password">
    '''   password required to log to the callback
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
    Public Function callbackLogin(ByVal username As String, ByVal password As String) As Integer
      Dim rest_val As String
      rest_val = username + ":" + password
      Return _setAttr("callbackCredentials", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the initial waiting time before first callback notifications, in seconds.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the initial waiting time before first callback notifications, in seconds
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_CALLBACKINITIALDELAY_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_callbackInitialDelay() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return CALLBACKINITIALDELAY_INVALID
        End If
      End If
      Return Me._callbackInitialDelay
    End Function


    '''*
    ''' <summary>
    '''   Changes the initial waiting time before first callback notifications, in seconds.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the initial waiting time before first callback notifications, in seconds
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
    Public Function set_callbackInitialDelay(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("callbackInitialDelay", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the minimum waiting time between two callback notifications, in seconds.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the minimum waiting time between two callback notifications, in seconds
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_CALLBACKMINDELAY_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_callbackMinDelay() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return CALLBACKMINDELAY_INVALID
        End If
      End If
      Return Me._callbackMinDelay
    End Function


    '''*
    ''' <summary>
    '''   Changes the minimum waiting time between two callback notifications, in seconds.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the minimum waiting time between two callback notifications, in seconds
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
    Public Function set_callbackMinDelay(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("callbackMinDelay", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the maximum waiting time between two callback notifications, in seconds.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the maximum waiting time between two callback notifications, in seconds
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_CALLBACKMAXDELAY_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_callbackMaxDelay() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return CALLBACKMAXDELAY_INVALID
        End If
      End If
      Return Me._callbackMaxDelay
    End Function


    '''*
    ''' <summary>
    '''   Changes the maximum waiting time between two callback notifications, in seconds.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the maximum waiting time between two callback notifications, in seconds
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
    Public Function set_callbackMaxDelay(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("callbackMaxDelay", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the current consumed by the module from Power-over-Ethernet (PoE), in milli-amps.
    ''' <para>
    '''   The current consumption is measured after converting PoE source to 5 Volt, and should
    '''   never exceed 1800 mA.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the current consumed by the module from Power-over-Ethernet (PoE), in milli-amps
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_POECURRENT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_poeCurrent() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return POECURRENT_INVALID
        End If
      End If
      Return Me._poeCurrent
    End Function

    '''*
    ''' <summary>
    '''   Retrieves a network interface for a given identifier.
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
    '''   This function does not require that the network interface is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YNetwork.isOnline()</c> to test if the network interface is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a network interface by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the network interface
    ''' </param>
    ''' <returns>
    '''   a <c>YNetwork</c> object allowing you to drive the network interface.
    ''' </returns>
    '''/
    Public Shared Function FindNetwork(func As String) As YNetwork
      Dim obj As YNetwork
      obj = CType(YFunction._FindFromCache("Network", func), YNetwork)
      If ((obj Is Nothing)) Then
        obj = New YNetwork(func)
        YFunction._AddToCache("Network", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YNetworkValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackNetwork = callback
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
      If (Not (Me._valueCallbackNetwork Is Nothing)) Then
        Me._valueCallbackNetwork(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Changes the configuration of the network interface to enable the use of an
    '''   IP address received from a DHCP server.
    ''' <para>
    '''   Until an address is received from a DHCP
    '''   server, the module uses the IP parameters specified to this function.
    '''   Remember to call the <c>saveToFlash()</c> method and then to reboot the module to apply this setting.
    ''' </para>
    ''' </summary>
    ''' <param name="fallbackIpAddr">
    '''   fallback IP address, to be used when no DHCP reply is received
    ''' </param>
    ''' <param name="fallbackSubnetMaskLen">
    '''   fallback subnet mask length when no DHCP reply is received, as an
    '''   integer (eg. 24 means 255.255.255.0)
    ''' </param>
    ''' <param name="fallbackRouter">
    '''   fallback router IP address, to be used when no DHCP reply is received
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function useDHCP(fallbackIpAddr As String, fallbackSubnetMaskLen As Integer, fallbackRouter As String) As Integer
      Return Me.set_ipConfig("DHCP:" +  fallbackIpAddr + "/" + Convert.ToString( fallbackSubnetMaskLen) + "/" + fallbackRouter)
    End Function

    '''*
    ''' <summary>
    '''   Changes the configuration of the network interface to use a static IP address.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method and then to reboot the module to apply this setting.
    ''' </para>
    ''' </summary>
    ''' <param name="ipAddress">
    '''   device IP address
    ''' </param>
    ''' <param name="subnetMaskLen">
    '''   subnet mask length, as an integer (eg. 24 means 255.255.255.0)
    ''' </param>
    ''' <param name="router">
    '''   router IP address (default gateway)
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function useStaticIP(ipAddress As String, subnetMaskLen As Integer, router As String) As Integer
      Return Me.set_ipConfig("STATIC:" +  ipAddress + "/" + Convert.ToString( subnetMaskLen) + "/" + router)
    End Function

    '''*
    ''' <summary>
    '''   Pings host to test the network connectivity.
    ''' <para>
    '''   Sends four ICMP ECHO_REQUEST requests from the
    '''   module to the target host. This method returns a string with the result of the
    '''   4 ICMP ECHO_REQUEST requests.
    ''' </para>
    ''' </summary>
    ''' <param name="host">
    '''   the hostname or the IP address of the target
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   a string with the result of the ping.
    ''' </returns>
    '''/
    Public Overridable Function ping(host As String) As String
      Dim content As Byte()
      REM // may throw an exception
      content = Me._download("ping.txt?host=" + host)
      Return YAPI.DefaultEncoding.GetString(content)
    End Function

    '''*
    ''' <summary>
    '''   Trigger an HTTP callback quickly.
    ''' <para>
    '''   This function can even be called within
    '''   an HTTP callback, in which case the next callback will be triggered 5 seconds
    '''   after the end of the current callback, regardless if the minimum time between
    '''   callbacks configured in the device.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function triggerCallback() As Integer
      REM // Rewrite the callback method to trigger the callback
      REM // may throw an exception
      Return Me.set_callbackMethod(Me.get_callbackMethod())
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of network interfaces started using <c>yFirstNetwork()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YNetwork</c> object, corresponding to
    '''   a network interface currently online, or a <c>Nothing</c> pointer
    '''   if there are no more network interfaces to enumerate.
    ''' </returns>
    '''/
    Public Function nextNetwork() As YNetwork
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YNetwork.FindNetwork(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of network interfaces currently accessible.
    ''' <para>
    '''   Use the method <c>YNetwork.nextNetwork()</c> to iterate on
    '''   next network interfaces.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YNetwork</c> object, corresponding to
    '''   the first network interface currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstNetwork() As YNetwork
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Network", 0, p, size, neededsize, errmsg)
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
      Return YNetwork.FindNetwork(serial + "." + funcId)
    End Function

    REM --- (end of YNetwork public methods declaration)

  End Class

  REM --- (Network functions)

  '''*
  ''' <summary>
  '''   Retrieves a network interface for a given identifier.
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
  '''   This function does not require that the network interface is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YNetwork.isOnline()</c> to test if the network interface is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a network interface by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the network interface
  ''' </param>
  ''' <returns>
  '''   a <c>YNetwork</c> object allowing you to drive the network interface.
  ''' </returns>
  '''/
  Public Function yFindNetwork(ByVal func As String) As YNetwork
    Return YNetwork.FindNetwork(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of network interfaces currently accessible.
  ''' <para>
  '''   Use the method <c>YNetwork.nextNetwork()</c> to iterate on
  '''   next network interfaces.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YNetwork</c> object, corresponding to
  '''   the first network interface currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstNetwork() As YNetwork
    Return YNetwork.FirstNetwork()
  End Function


  REM --- (end of Network functions)

End Module

'*********************************************************************
'*
'* $Id: yocto_wireless.vb 9921 2013-02-20 09:39:16Z seb $
'*
'* Implements yFindWireless(), the high-level API for Wireless functions
'*
'* - - - - - - - - - License information: - - - - - - - - - 
'*
'* Copyright (C) 2011 and beyond by Yoctopuce Sarl, Switzerland.
'*
'* 1) If you have obtained this file from www.yoctopuce.com,
'*    Yoctopuce Sarl licenses to you (hereafter Licensee) the
'*    right to use, modify, copy, and integrate this source file
'*    into your own solution for the sole purpose of interfacing
'*    a Yoctopuce product with Licensee's solution.
'*
'*    The use of this file and all relationship between Yoctopuce 
'*    and Licensee are governed by Yoctopuce General Terms and 
'*    Conditions.
'*
'*    THE SOFTWARE AND DOCUMENTATION ARE PROVIDED 'AS IS' WITHOUT
'*    WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING 
'*    WITHOUT LIMITATION, ANY WARRANTY OF MERCHANTABILITY, FITNESS 
'*    FOR A PARTICULAR PURPOSE, TITLE AND NON-INFRINGEMENT. IN NO
'*    EVENT SHALL LICENSOR BE LIABLE FOR ANY INCIDENTAL, SPECIAL,
'*    INDIRECT OR CONSEQUENTIAL DAMAGES, LOST PROFITS OR LOST DATA, 
'*    COST OF PROCUREMENT OF SUBSTITUTE GOODS, TECHNOLOGY OR 
'*    SERVICES, ANY CLAIMS BY THIRD PARTIES (INCLUDING BUT NOT 
'*    LIMITED TO ANY DEFENSE THEREOF), ANY CLAIMS FOR INDEMNITY OR
'*    CONTRIBUTION, OR OTHER SIMILAR COSTS, WHETHER ASSERTED ON THE
'*    BASIS OF CONTRACT, TORT (INCLUDING NEGLIGENCE), BREACH OF
'*    WARRANTY, OR OTHERWISE.
'*
'* 2) If your intent is not to interface with Yoctopuce products,
'*    you are not entitled to use, read or create any derived
'*    material from this source file.
'*
'*********************************************************************/


Imports YDEV_DESCR = System.Int32
Imports YFUN_DESCR = System.Int32
Imports System.Runtime.InteropServices
Imports System.Text

Module yocto_wireless

  REM --- (YWireless definitions)

  Public Delegate Sub UpdateCallback(ByVal func As YWireless, ByVal value As String)


  Public Const Y_LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_LINKQUALITY_INVALID As Integer = YAPI.INVALID_UNSIGNED
  Public Const Y_SSID_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_CHANNEL_INVALID As Integer = YAPI.INVALID_UNSIGNED
  Public Const Y_SECURITY_UNKNOWN = 0
  Public Const Y_SECURITY_OPEN = 1
  Public Const Y_SECURITY_WEP = 2
  Public Const Y_SECURITY_WPA = 3
  Public Const Y_SECURITY_WPA2 = 4
  Public Const Y_SECURITY_INVALID = -1

  Public Const Y_MESSAGE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_WLANCONFIG_INVALID As String = YAPI.INVALID_STRING


  REM --- (end of YWireless definitions)

  REM --- (YWireless implementation)

  Private _WirelessCache As New Hashtable()
  Private _callback As UpdateCallback

  Public Class YWireless
    Inherits YFunction
    Public Const LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
    Public Const ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
    Public Const LINKQUALITY_INVALID As Integer = YAPI.INVALID_UNSIGNED
    Public Const SSID_INVALID As String = YAPI.INVALID_STRING
    Public Const CHANNEL_INVALID As Integer = YAPI.INVALID_UNSIGNED
    Public Const SECURITY_UNKNOWN = 0
    Public Const SECURITY_OPEN = 1
    Public Const SECURITY_WEP = 2
    Public Const SECURITY_WPA = 3
    Public Const SECURITY_WPA2 = 4
    Public Const SECURITY_INVALID = -1

    Public Const MESSAGE_INVALID As String = YAPI.INVALID_STRING
    Public Const WLANCONFIG_INVALID As String = YAPI.INVALID_STRING

    Protected _logicalName As String
    Protected _advertisedValue As String
    Protected _linkQuality As Long
    Protected _ssid As String
    Protected _channel As Long
    Protected _security As Long
    Protected _message As String
    Protected _wlanConfig As String

    Public Sub New(ByVal func As String)
      MyBase.new("Wireless", func)
      _logicalName = Y_LOGICALNAME_INVALID
      _advertisedValue = Y_ADVERTISEDVALUE_INVALID
      _linkQuality = Y_LINKQUALITY_INVALID
      _ssid = Y_SSID_INVALID
      _channel = Y_CHANNEL_INVALID
      _security = Y_SECURITY_INVALID
      _message = Y_MESSAGE_INVALID
      _wlanConfig = Y_WLANCONFIG_INVALID
    End Sub

    Protected Overrides Function _parse(ByRef j As TJSONRECORD) As Integer
      Dim member As TJSONRECORD
      Dim i As Integer
      If (j.recordtype <> TJSONRECORDTYPE.JSON_STRUCT) Then
        Return -1
      End If
      For i = 0 To j.membercount - 1
        member = j.members(i)
        If (member.name = "logicalName") Then
          _logicalName = member.svalue
        ElseIf (member.name = "advertisedValue") Then
          _advertisedValue = member.svalue
        ElseIf (member.name = "linkQuality") Then
          _linkQuality = CLng(member.ivalue)
        ElseIf (member.name = "ssid") Then
          _ssid = member.svalue
        ElseIf (member.name = "channel") Then
          _channel = CLng(member.ivalue)
        ElseIf (member.name = "security") Then
          _security = CLng(member.ivalue)
        ElseIf (member.name = "message") Then
          _message = member.svalue
        ElseIf (member.name = "wlanConfig") Then
          _wlanConfig = member.svalue
        End If
      Next i
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Returns the logical name of the wireless lan interface.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the logical name of the wireless lan interface
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_LOGICALNAME_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_logicalName() As String
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_LOGICALNAME_INVALID
        End If
      End If
      Return _logicalName
    End Function

    '''*
    ''' <summary>
    '''   Changes the logical name of the wireless lan interface.
    ''' <para>
    '''   You can use <c>yCheckLogicalName()</c>
    '''   prior to this call to make sure that your parameter is valid.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the logical name of the wireless lan interface
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
    Public Function set_logicalName(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("logicalName", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the current value of the wireless lan interface (no more than 6 characters).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the current value of the wireless lan interface (no more than 6 characters)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_ADVERTISEDVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_advertisedValue() As String
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_ADVERTISEDVALUE_INVALID
        End If
      End If
      Return _advertisedValue
    End Function

    '''*
    ''' <summary>
    '''   Returns the link quality, expressed in per cents.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the link quality, expressed in per cents
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_LINKQUALITY_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_linkQuality() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_LINKQUALITY_INVALID
        End If
      End If
      Return CType(_linkQuality,Integer)
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
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_SSID_INVALID
        End If
      End If
      Return _ssid
    End Function

    '''*
    ''' <summary>
    '''   Returns the 802.
    ''' <para>
    '''   11 channel currently used, or 0 when the selected network has not been found.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the 802
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_CHANNEL_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_channel() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_CHANNEL_INVALID
        End If
      End If
      Return CType(_channel,Integer)
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
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_SECURITY_INVALID
        End If
      End If
      Return CType(_security,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Returns the last status message from the wireless interface.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the last status message from the wireless interface
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_MESSAGE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_message() As String
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_MESSAGE_INVALID
        End If
      End If
      Return _message
    End Function

    Public Function get_wlanConfig() As String
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_WLANCONFIG_INVALID
        End If
      End If
      Return _wlanConfig
    End Function

    Public Function set_wlanConfig(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("wlanConfig", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Changes the configuration of the wireless lan interface to connect to an existing
    '''   access point (infrastructure mode).
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method and then to reboot the module to apply this setting.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="ssid">
    '''   the name of the network to connect to
    ''' </param>
    ''' <param name="securityKey">
    '''   the network key, as a character string
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
    Public Function joinNetwork(ByVal ssid As String,ByVal securityKey As String) As Integer
      Dim rest_val As String
      rest_val = "INFRA:"+ssid+"\\"+securityKey
      Return _setAttr("wlanConfig", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Changes the configuration of the wireless lan interface to create an ad-hoc
    '''   wireless network, without using an access point.
    ''' <para>
    '''   If a security key is specified,
    '''   the network will be protected by WEP128, since WPA is not standardized for
    '''   ad-hoc networks.
    '''   Remember to call the <c>saveToFlash()</c> method and then to reboot the module to apply this setting.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="ssid">
    '''   the name of the network to connect to
    ''' </param>
    ''' <param name="securityKey">
    '''   the network key, as a character string
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
    Public Function adhocNetwork(ByVal ssid As String,ByVal securityKey As String) As Integer
      Dim rest_val As String
      rest_val = "ADHOC:"+ssid+"\\"+securityKey
      Return _setAttr("wlanConfig", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Continues the enumeration of wireless lan interfaces started using <c>yFirstWireless()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YWireless</c> object, corresponding to
    '''   a wireless lan interface currently online, or a <c>null</c> pointer
    '''   if there are no more wireless lan interfaces to enumerate.
    ''' </returns>
    '''/
    Public Function nextWireless() as YWireless
      Dim hwid As String =""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid="") Then
        Return Nothing
      End If
      Return yFindWireless(hwid)
    End Function

    '''*
    ''' <summary>
    '''   comment from .
    ''' <para>
    '''   yc definition
    ''' </para>
    ''' </summary>
    '''/
  Public Overloads Sub registerValueCallback(ByVal callback As UpdateCallback)
   If (callback IsNot Nothing) Then
     registerFuncCallback(Me)
   Else
     unregisterFuncCallback(Me)
   End If
   _callback = callback
  End Sub

  Public Sub set_callback(ByVal callback As UpdateCallback)
    registerValueCallback(callback)
  End Sub

  Public Sub setCallback(ByVal callback As UpdateCallback)
    registerValueCallback(callback)
  End Sub

  Public Overrides Sub advertiseValue(ByVal value As String)
    If (_callback IsNot Nothing) Then _callback(Me, value)
  End Sub


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
    Public Shared Function FindWireless(ByVal func As String) As YWireless
      Dim res As YWireless
      If (_WirelessCache.ContainsKey(func)) Then
        Return CType(_WirelessCache(func), YWireless)
      End If
      res = New YWireless(func)
      _WirelessCache.Add(func, res)
      Return res
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
    '''   the first wireless lan interface currently online, or a <c>null</c> pointer
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

    REM --- (end of YWireless implementation)

  End Class

  REM --- (Wireless functions)

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
  '''   the first wireless lan interface currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstWireless() As YWireless
    Return YWireless.FirstWireless()
  End Function

  Private Sub _WirelessCleanup()
  End Sub


  REM --- (end of Wireless functions)

End Module

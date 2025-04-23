' ********************************************************************
'
'  $Id: svn_id $
'
'  Implements yFindSpectralChannel(), the high-level API for SpectralChannel functions
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

Module yocto_spectralchannel

    REM --- (YSpectralChannel return codes)
    REM --- (end of YSpectralChannel return codes)
    REM --- (YSpectralChannel dlldef)
    REM --- (end of YSpectralChannel dlldef)
   REM --- (YSpectralChannel yapiwrapper)
   REM --- (end of YSpectralChannel yapiwrapper)
  REM --- (YSpectralChannel globals)

  Public Const Y_RAWCOUNT_INVALID As Integer = YAPI.INVALID_INT
  Public Const Y_CHANNELNAME_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_PEAKWAVELENGTH_INVALID As Integer = YAPI.INVALID_INT
  Public Delegate Sub YSpectralChannelValueCallback(ByVal func As YSpectralChannel, ByVal value As String)
  Public Delegate Sub YSpectralChannelTimedReportCallback(ByVal func As YSpectralChannel, ByVal measure As YMeasure)
  REM --- (end of YSpectralChannel globals)

  REM --- (YSpectralChannel class start)

  '''*
  ''' <summary>
  '''   The <c>YSpectralChannel</c> class allows you to read and configure Yoctopuce spectral analysis channels.
  ''' <para>
  '''   It inherits from <c>YSensor</c> class the core functions to read measures,
  '''   to register callback functions, and to access the autonomous datalogger.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YSpectralChannel
    Inherits YSensor
    REM --- (end of YSpectralChannel class start)

    REM --- (YSpectralChannel definitions)
    Public Const RAWCOUNT_INVALID As Integer = YAPI.INVALID_INT
    Public Const CHANNELNAME_INVALID As String = YAPI.INVALID_STRING
    Public Const PEAKWAVELENGTH_INVALID As Integer = YAPI.INVALID_INT
    REM --- (end of YSpectralChannel definitions)

    REM --- (YSpectralChannel attributes declaration)
    Protected _rawCount As Integer
    Protected _channelName As String
    Protected _peakWavelength As Integer
    Protected _valueCallbackSpectralChannel As YSpectralChannelValueCallback
    Protected _timedReportCallbackSpectralChannel As YSpectralChannelTimedReportCallback
    REM --- (end of YSpectralChannel attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "SpectralChannel"
      REM --- (YSpectralChannel attributes initialization)
      _rawCount = RAWCOUNT_INVALID
      _channelName = CHANNELNAME_INVALID
      _peakWavelength = PEAKWAVELENGTH_INVALID
      _valueCallbackSpectralChannel = Nothing
      _timedReportCallbackSpectralChannel = Nothing
      REM --- (end of YSpectralChannel attributes initialization)
    End Sub

    REM --- (YSpectralChannel private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("rawCount") Then
        _rawCount = CInt(json_val.getLong("rawCount"))
      End If
      If json_val.has("channelName") Then
        _channelName = json_val.getString("channelName")
      End If
      If json_val.has("peakWavelength") Then
        _peakWavelength = CInt(json_val.getLong("peakWavelength"))
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YSpectralChannel private methods declaration)

    REM --- (YSpectralChannel public methods declaration)
    '''*
    ''' <summary>
    '''   Retrieves the raw spectral intensity value as measured by the sensor, without any scaling or calibration.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSpectralChannel.RAWCOUNT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_rawCount() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return RAWCOUNT_INVALID
        End If
      End If
      res = Me._rawCount
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the target spectral band name.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the target spectral band name
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSpectralChannel.CHANNELNAME_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_channelName() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return CHANNELNAME_INVALID
        End If
      End If
      res = Me._channelName
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the target spectral band peak wavelenght, in nm.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the target spectral band peak wavelenght, in nm
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSpectralChannel.PEAKWAVELENGTH_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_peakWavelength() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return PEAKWAVELENGTH_INVALID
        End If
      End If
      res = Me._peakWavelength
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Retrieves a spectral analysis channel for a given identifier.
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
    '''   This function does not require that the spectral analysis channel is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YSpectralChannel.isOnline()</c> to test if the spectral analysis channel is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a spectral analysis channel by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the spectral analysis channel, for instance
    '''   <c>MyDevice.spectralChannel1</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YSpectralChannel</c> object allowing you to drive the spectral analysis channel.
    ''' </returns>
    '''/
    Public Shared Function FindSpectralChannel(func As String) As YSpectralChannel
      Dim obj As YSpectralChannel
      obj = CType(YFunction._FindFromCache("SpectralChannel", func), YSpectralChannel)
      If ((obj Is Nothing)) Then
        obj = New YSpectralChannel(func)
        YFunction._AddToCache("SpectralChannel", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YSpectralChannelValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackSpectralChannel = callback
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
      If (Not (Me._valueCallbackSpectralChannel Is Nothing)) Then
        Me._valueCallbackSpectralChannel(Me, value)
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
    Public Overloads Function registerTimedReportCallback(callback As YSpectralChannelTimedReportCallback) As Integer
      Dim sensor As YSensor
      sensor = Me
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(sensor, True)
      Else
        YFunction._UpdateTimedReportCallbackList(sensor, False)
      End If
      Me._timedReportCallbackSpectralChannel = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackSpectralChannel Is Nothing)) Then
        Me._timedReportCallbackSpectralChannel(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of spectral analysis channels started using <c>yFirstSpectralChannel()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned spectral analysis channels order.
    '''   If you want to find a specific a spectral analysis channel, use <c>SpectralChannel.findSpectralChannel()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YSpectralChannel</c> object, corresponding to
    '''   a spectral analysis channel currently online, or a <c>Nothing</c> pointer
    '''   if there are no more spectral analysis channels to enumerate.
    ''' </returns>
    '''/
    Public Function nextSpectralChannel() As YSpectralChannel
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YSpectralChannel.FindSpectralChannel(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of spectral analysis channels currently accessible.
    ''' <para>
    '''   Use the method <c>YSpectralChannel.nextSpectralChannel()</c> to iterate on
    '''   next spectral analysis channels.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YSpectralChannel</c> object, corresponding to
    '''   the first spectral analysis channel currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstSpectralChannel() As YSpectralChannel
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("SpectralChannel", 0, p, size, neededsize, errmsg)
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
      Return YSpectralChannel.FindSpectralChannel(serial + "." + funcId)
    End Function

    REM --- (end of YSpectralChannel public methods declaration)

  End Class

  REM --- (YSpectralChannel functions)

  '''*
  ''' <summary>
  '''   Retrieves a spectral analysis channel for a given identifier.
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
  '''   This function does not require that the spectral analysis channel is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YSpectralChannel.isOnline()</c> to test if the spectral analysis channel is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a spectral analysis channel by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the spectral analysis channel, for instance
  '''   <c>MyDevice.spectralChannel1</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YSpectralChannel</c> object allowing you to drive the spectral analysis channel.
  ''' </returns>
  '''/
  Public Function yFindSpectralChannel(ByVal func As String) As YSpectralChannel
    Return YSpectralChannel.FindSpectralChannel(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of spectral analysis channels currently accessible.
  ''' <para>
  '''   Use the method <c>YSpectralChannel.nextSpectralChannel()</c> to iterate on
  '''   next spectral analysis channels.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YSpectralChannel</c> object, corresponding to
  '''   the first spectral analysis channel currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstSpectralChannel() As YSpectralChannel
    Return YSpectralChannel.FirstSpectralChannel()
  End Function


  REM --- (end of YSpectralChannel functions)

End Module

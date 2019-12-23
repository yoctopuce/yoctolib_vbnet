' ********************************************************************
'
'  $Id: yocto_audioout.vb 38899 2019-12-20 17:21:03Z mvuilleu $
'
'  Implements yFindAudioOut(), the high-level API for AudioOut functions
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

Module yocto_audioout

    REM --- (YAudioOut return codes)
    REM --- (end of YAudioOut return codes)
    REM --- (YAudioOut dlldef)
    REM --- (end of YAudioOut dlldef)
   REM --- (YAudioOut yapiwrapper)
   REM --- (end of YAudioOut yapiwrapper)
  REM --- (YAudioOut globals)

  Public Const Y_VOLUME_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_MUTE_FALSE As Integer = 0
  Public Const Y_MUTE_TRUE As Integer = 1
  Public Const Y_MUTE_INVALID As Integer = -1
  Public Const Y_VOLUMERANGE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_SIGNAL_INVALID As Integer = YAPI.INVALID_INT
  Public Const Y_NOSIGNALFOR_INVALID As Integer = YAPI.INVALID_INT
  Public Delegate Sub YAudioOutValueCallback(ByVal func As YAudioOut, ByVal value As String)
  Public Delegate Sub YAudioOutTimedReportCallback(ByVal func As YAudioOut, ByVal measure As YMeasure)
  REM --- (end of YAudioOut globals)

  REM --- (YAudioOut class start)

  '''*
  ''' <summary>
  '''   The <c>YAudioOut</c> class allows you to configure the volume of an audio output.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YAudioOut
    Inherits YFunction
    REM --- (end of YAudioOut class start)

    REM --- (YAudioOut definitions)
    Public Const VOLUME_INVALID As Integer = YAPI.INVALID_UINT
    Public Const MUTE_FALSE As Integer = 0
    Public Const MUTE_TRUE As Integer = 1
    Public Const MUTE_INVALID As Integer = -1
    Public Const VOLUMERANGE_INVALID As String = YAPI.INVALID_STRING
    Public Const SIGNAL_INVALID As Integer = YAPI.INVALID_INT
    Public Const NOSIGNALFOR_INVALID As Integer = YAPI.INVALID_INT
    REM --- (end of YAudioOut definitions)

    REM --- (YAudioOut attributes declaration)
    Protected _volume As Integer
    Protected _mute As Integer
    Protected _volumeRange As String
    Protected _signal As Integer
    Protected _noSignalFor As Integer
    Protected _valueCallbackAudioOut As YAudioOutValueCallback
    REM --- (end of YAudioOut attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "AudioOut"
      REM --- (YAudioOut attributes initialization)
      _volume = VOLUME_INVALID
      _mute = MUTE_INVALID
      _volumeRange = VOLUMERANGE_INVALID
      _signal = SIGNAL_INVALID
      _noSignalFor = NOSIGNALFOR_INVALID
      _valueCallbackAudioOut = Nothing
      REM --- (end of YAudioOut attributes initialization)
    End Sub

    REM --- (YAudioOut private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("volume") Then
        _volume = CInt(json_val.getLong("volume"))
      End If
      If json_val.has("mute") Then
        If (json_val.getInt("mute") > 0) Then _mute = 1 Else _mute = 0
      End If
      If json_val.has("volumeRange") Then
        _volumeRange = json_val.getString("volumeRange")
      End If
      If json_val.has("signal") Then
        _signal = CInt(json_val.getLong("signal"))
      End If
      If json_val.has("noSignalFor") Then
        _noSignalFor = CInt(json_val.getLong("noSignalFor"))
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YAudioOut private methods declaration)

    REM --- (YAudioOut public methods declaration)
    '''*
    ''' <summary>
    '''   Returns audio output volume, in per cents.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to audio output volume, in per cents
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
    '''   Changes audio output volume, in per cents.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c>
    '''   method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to audio output volume, in per cents
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
    '''   Returns the supported volume range.
    ''' <para>
    '''   The low value of the
    '''   range corresponds to the minimal audible value. To
    '''   completely mute the sound, use <c>set_mute()</c>
    '''   instead of the <c>set_volume()</c>.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the supported volume range
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_VOLUMERANGE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_volumeRange() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return VOLUMERANGE_INVALID
        End If
      End If
      res = Me._volumeRange
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the detected output current level.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the detected output current level
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_SIGNAL_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_signal() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return SIGNAL_INVALID
        End If
      End If
      res = Me._signal
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the number of seconds elapsed without detecting a signal.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of seconds elapsed without detecting a signal
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_NOSIGNALFOR_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_noSignalFor() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return NOSIGNALFOR_INVALID
        End If
      End If
      res = Me._noSignalFor
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Retrieves an audio output for a given identifier.
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
    '''   This function does not require that the audio output is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YAudioOut.isOnline()</c> to test if the audio output is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   an audio output by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the audio output, for instance
    '''   <c>MyDevice.audioOut1</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YAudioOut</c> object allowing you to drive the audio output.
    ''' </returns>
    '''/
    Public Shared Function FindAudioOut(func As String) As YAudioOut
      Dim obj As YAudioOut
      obj = CType(YFunction._FindFromCache("AudioOut", func), YAudioOut)
      If ((obj Is Nothing)) Then
        obj = New YAudioOut(func)
        YFunction._AddToCache("AudioOut", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YAudioOutValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackAudioOut = callback
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
      If (Not (Me._valueCallbackAudioOut Is Nothing)) Then
        Me._valueCallbackAudioOut(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of audio outputs started using <c>yFirstAudioOut()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned audio outputs order.
    '''   If you want to find a specific an audio output, use <c>AudioOut.findAudioOut()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YAudioOut</c> object, corresponding to
    '''   an audio output currently online, or a <c>Nothing</c> pointer
    '''   if there are no more audio outputs to enumerate.
    ''' </returns>
    '''/
    Public Function nextAudioOut() As YAudioOut
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YAudioOut.FindAudioOut(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of audio outputs currently accessible.
    ''' <para>
    '''   Use the method <c>YAudioOut.nextAudioOut()</c> to iterate on
    '''   next audio outputs.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YAudioOut</c> object, corresponding to
    '''   the first audio output currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstAudioOut() As YAudioOut
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("AudioOut", 0, p, size, neededsize, errmsg)
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
      Return YAudioOut.FindAudioOut(serial + "." + funcId)
    End Function

    REM --- (end of YAudioOut public methods declaration)

  End Class

  REM --- (YAudioOut functions)

  '''*
  ''' <summary>
  '''   Retrieves an audio output for a given identifier.
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
  '''   This function does not require that the audio output is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YAudioOut.isOnline()</c> to test if the audio output is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   an audio output by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the audio output, for instance
  '''   <c>MyDevice.audioOut1</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YAudioOut</c> object allowing you to drive the audio output.
  ''' </returns>
  '''/
  Public Function yFindAudioOut(ByVal func As String) As YAudioOut
    Return YAudioOut.FindAudioOut(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of audio outputs currently accessible.
  ''' <para>
  '''   Use the method <c>YAudioOut.nextAudioOut()</c> to iterate on
  '''   next audio outputs.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YAudioOut</c> object, corresponding to
  '''   the first audio output currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstAudioOut() As YAudioOut
    Return YAudioOut.FirstAudioOut()
  End Function


  REM --- (end of YAudioOut functions)

End Module

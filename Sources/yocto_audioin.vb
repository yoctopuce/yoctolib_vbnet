'*********************************************************************
'*
'* $Id: yocto_audioin.vb 20797 2015-07-06 16:49:40Z mvuilleu $
'*
'* Implements yFindAudioIn(), the high-level API for AudioIn functions
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

Module yocto_audioin

    REM --- (YAudioIn return codes)
    REM --- (end of YAudioIn return codes)
    REM --- (YAudioIn dlldef)
    REM --- (end of YAudioIn dlldef)
  REM --- (YAudioIn globals)

  Public Const Y_VOLUME_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_MUTE_FALSE As Integer = 0
  Public Const Y_MUTE_TRUE As Integer = 1
  Public Const Y_MUTE_INVALID As Integer = -1
  Public Const Y_VOLUMERANGE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_SIGNAL_INVALID As Integer = YAPI.INVALID_INT
  Public Const Y_NOSIGNALFOR_INVALID As Integer = YAPI.INVALID_INT
  Public Delegate Sub YAudioInValueCallback(ByVal func As YAudioIn, ByVal value As String)
  Public Delegate Sub YAudioInTimedReportCallback(ByVal func As YAudioIn, ByVal measure As YMeasure)
  REM --- (end of YAudioIn globals)

  REM --- (YAudioIn class start)

  '''*
  ''' <summary>
  '''   The Yoctopuce application programming interface allows you to configure the volume of the input channel.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YAudioIn
    Inherits YFunction
    REM --- (end of YAudioIn class start)

    REM --- (YAudioIn definitions)
    Public Const VOLUME_INVALID As Integer = YAPI.INVALID_UINT
    Public Const MUTE_FALSE As Integer = 0
    Public Const MUTE_TRUE As Integer = 1
    Public Const MUTE_INVALID As Integer = -1
    Public Const VOLUMERANGE_INVALID As String = YAPI.INVALID_STRING
    Public Const SIGNAL_INVALID As Integer = YAPI.INVALID_INT
    Public Const NOSIGNALFOR_INVALID As Integer = YAPI.INVALID_INT
    REM --- (end of YAudioIn definitions)

    REM --- (YAudioIn attributes declaration)
    Protected _volume As Integer
    Protected _mute As Integer
    Protected _volumeRange As String
    Protected _signal As Integer
    Protected _noSignalFor As Integer
    Protected _valueCallbackAudioIn As YAudioInValueCallback
    REM --- (end of YAudioIn attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "AudioIn"
      REM --- (YAudioIn attributes initialization)
      _volume = VOLUME_INVALID
      _mute = MUTE_INVALID
      _volumeRange = VOLUMERANGE_INVALID
      _signal = SIGNAL_INVALID
      _noSignalFor = NOSIGNALFOR_INVALID
      _valueCallbackAudioIn = Nothing
      REM --- (end of YAudioIn attributes initialization)
    End Sub

    REM --- (YAudioIn private methods declaration)

    Protected Overrides Function _parseAttr(ByRef member As TJSONRECORD) As Integer
      If (member.name = "volume") Then
        _volume = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "mute") Then
        If (member.ivalue > 0) Then _mute = 1 Else _mute = 0
        Return 1
      End If
      If (member.name = "volumeRange") Then
        _volumeRange = member.svalue
        Return 1
      End If
      If (member.name = "signal") Then
        _signal = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "noSignalFor") Then
        _noSignalFor = CInt(member.ivalue)
        Return 1
      End If
      Return MyBase._parseAttr(member)
    End Function

    REM --- (end of YAudioIn private methods declaration)

    REM --- (YAudioIn public methods declaration)
    '''*
    ''' <summary>
    '''   Returns audio input gain, in per cents.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to audio input gain, in per cents
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_VOLUME_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_volume() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return VOLUME_INVALID
        End If
      End If
      Return Me._volume
    End Function


    '''*
    ''' <summary>
    '''   Changes audio input gain, in per cents.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to audio input gain, in per cents
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
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return MUTE_INVALID
        End If
      End If
      Return Me._mute
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
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return VOLUMERANGE_INVALID
        End If
      End If
      Return Me._volumeRange
    End Function

    '''*
    ''' <summary>
    '''   Returns the detected input signal level.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the detected input signal level
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_SIGNAL_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_signal() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return SIGNAL_INVALID
        End If
      End If
      Return Me._signal
    End Function

    '''*
    ''' <summary>
    '''   Returns the number of seconds elapsed without detecting a signal
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
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return NOSIGNALFOR_INVALID
        End If
      End If
      Return Me._noSignalFor
    End Function

    '''*
    ''' <summary>
    '''   Retrieves an audio input for a given identifier.
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
    '''   This function does not require that the audio input is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YAudioIn.isOnline()</c> to test if the audio input is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   an audio input by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the audio input
    ''' </param>
    ''' <returns>
    '''   a <c>YAudioIn</c> object allowing you to drive the audio input.
    ''' </returns>
    '''/
    Public Shared Function FindAudioIn(func As String) As YAudioIn
      Dim obj As YAudioIn
      obj = CType(YFunction._FindFromCache("AudioIn", func), YAudioIn)
      If ((obj Is Nothing)) Then
        obj = New YAudioIn(func)
        YFunction._AddToCache("AudioIn", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YAudioInValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackAudioIn = callback
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
      If (Not (Me._valueCallbackAudioIn Is Nothing)) Then
        Me._valueCallbackAudioIn(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of audio inputs started using <c>yFirstAudioIn()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YAudioIn</c> object, corresponding to
    '''   an audio input currently online, or a <c>null</c> pointer
    '''   if there are no more audio inputs to enumerate.
    ''' </returns>
    '''/
    Public Function nextAudioIn() As YAudioIn
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YAudioIn.FindAudioIn(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of audio inputs currently accessible.
    ''' <para>
    '''   Use the method <c>YAudioIn.nextAudioIn()</c> to iterate on
    '''   next audio inputs.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YAudioIn</c> object, corresponding to
    '''   the first audio input currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstAudioIn() As YAudioIn
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("AudioIn", 0, p, size, neededsize, errmsg)
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
      Return YAudioIn.FindAudioIn(serial + "." + funcId)
    End Function

    REM --- (end of YAudioIn public methods declaration)

  End Class

  REM --- (AudioIn functions)

  '''*
  ''' <summary>
  '''   Retrieves an audio input for a given identifier.
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
  '''   This function does not require that the audio input is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YAudioIn.isOnline()</c> to test if the audio input is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   an audio input by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the audio input
  ''' </param>
  ''' <returns>
  '''   a <c>YAudioIn</c> object allowing you to drive the audio input.
  ''' </returns>
  '''/
  Public Function yFindAudioIn(ByVal func As String) As YAudioIn
    Return YAudioIn.FindAudioIn(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of audio inputs currently accessible.
  ''' <para>
  '''   Use the method <c>YAudioIn.nextAudioIn()</c> to iterate on
  '''   next audio inputs.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YAudioIn</c> object, corresponding to
  '''   the first audio input currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstAudioIn() As YAudioIn
    Return YAudioIn.FirstAudioIn()
  End Function


  REM --- (end of AudioIn functions)

End Module
